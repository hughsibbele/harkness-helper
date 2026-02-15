/**
 * Canvas LMS API integration
 *
 * Handles student roster sync, grade posting, and course data fetching.
 * Supports both group and individual grading modes.
 */

// ============================================================================
// CANVAS API CORE
// ============================================================================

/**
 * Make an authenticated request to Canvas API
 * @param {string} endpoint - API endpoint (without base URL)
 * @param {string} method - HTTP method (GET, POST, PUT, DELETE)
 * @param {Object} payload - Request body for POST/PUT
 * @returns {Object} Parsed JSON response
 */
function canvasRequest(endpoint, method = 'GET', payload = null) {
  const baseUrl = getCanvasBaseUrl();
  const token = getCanvasToken();

  const url = `${baseUrl}/api/v1${endpoint}`;

  const options = {
    method: method,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload && (method === 'POST' || method === 'PUT')) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code >= 400) {
    const errorText = response.getContentText();
    Logger.log(`Canvas API error (${code}): ${errorText}`);
    throw new Error(`Canvas API error ${code}: ${errorText}`);
  }

  const contentText = response.getContentText();
  return contentText ? JSON.parse(contentText) : {};
}

/**
 * Handle paginated Canvas API responses
 * @param {string} endpoint - API endpoint
 * @returns {Object[]} All results across pages
 */
function canvasRequestPaginated(endpoint) {
  const baseUrl = getCanvasBaseUrl();
  const token = getCanvasToken();

  let allResults = [];
  let url = `${baseUrl}/api/v1${endpoint}`;

  if (!url.includes('per_page=')) {
    url += (url.includes('?') ? '&' : '?') + 'per_page=100';
  }

  while (url) {
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() >= 400) {
      throw new Error(`Canvas API error: ${response.getContentText()}`);
    }

    const results = JSON.parse(response.getContentText());
    allResults = allResults.concat(results);

    const linkHeader = response.getHeaders()['Link'] || '';
    const nextMatch = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
    url = nextMatch ? nextMatch[1] : null;
  }

  return allResults;
}

// ============================================================================
// COURSE & USER OPERATIONS
// ============================================================================

/**
 * Get students enrolled in a course
 * @param {string} courseId
 * @returns {Object[]} Array of student objects
 */
function getCanvasStudents(courseId) {
  const students = canvasRequestPaginated(
    `/courses/${courseId}/users?enrollment_type[]=student&include[]=email`
  );

  return students.map(s => ({
    canvasUserId: s.id,
    name: s.name,
    sortableName: s.sortable_name,
    email: s.email || '',
    loginId: s.login_id || ''
  }));
}

/**
 * Sync Canvas students to local Students sheet
 * @param {string} courseId
 * @param {string} classPeriod - Class period to assign
 * @returns {number} Number of students synced
 */
function syncCanvasStudents(courseId, classPeriod) {
  const canvasStudents = getCanvasStudents(courseId);

  for (const student of canvasStudents) {
    upsertStudent({
      name: student.name,
      email: student.email,
      class_period: classPeriod,
      canvas_user_id: student.canvasUserId.toString()
    });
  }

  Logger.log(`Synced ${canvasStudents.length} students from Canvas course ${courseId}`);
  return canvasStudents.length;
}

/**
 * Strip HTML tags from a string
 * @param {string} html
 * @returns {string} Plain text
 */
function stripHtml(html) {
  if (!html) return '';

  let text = html.replace(/<[^>]*>/g, ' ');

  text = text.replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");

  text = text.replace(/\s+/g, ' ').trim();

  return text;
}

// ============================================================================
// GRADE POSTING
// ============================================================================

/**
 * Post a grade and comment to a Canvas assignment
 * @param {string} courseId
 * @param {string} assignmentId
 * @param {string} studentUserId - Canvas user ID
 * @param {number|string} grade - The grade to post
 * @param {string} comment - Optional comment/feedback
 * @returns {Object} Submission result
 */
function postGrade(courseId, assignmentId, studentUserId, grade, comment = null) {
  const payload = {
    submission: {
      posted_grade: grade.toString()
    }
  };

  if (comment) {
    payload.comment = {
      text_comment: comment
    };
  }

  const result = canvasRequest(
    `/courses/${courseId}/assignments/${assignmentId}/submissions/${studentUserId}`,
    'PUT',
    payload
  );

  Logger.log(`Posted grade ${grade} for user ${studentUserId} on assignment ${assignmentId}`);
  return result;
}

/**
 * Post grades for a discussion to Canvas.
 * Branches by mode: group posts same grade to all, individual posts per-student.
 *
 * @param {string} discussionId
 * @returns {Object} {success, failed, errors}
 */
function postGradesForDiscussion(discussionId) {
  const discussion = getDiscussion(discussionId);
  if (!discussion) {
    throw new Error(`Discussion not found: ${discussionId}`);
  }

  if (!discussion.canvas_assignment_id) {
    throw new Error('No Canvas assignment configured for this discussion');
  }

  const courseId = getSetting('canvas_course_id');
  if (!courseId) {
    throw new Error('Canvas course ID not set in Settings');
  }

  const results = { success: 0, failed: 0, errors: [] };

  if (isGroupMode()) {
    // Group mode: post discussion.grade + discussion.group_feedback to all students
    const grade = discussion.grade;
    const feedback = discussion.group_feedback || '';

    if (!grade) {
      throw new Error('No grade set on discussion row');
    }

    const students = discussion.class_period
      ? getStudentsByClass(discussion.class_period)
      : getAllRows(CONFIG.SHEETS.STUDENTS);

    for (const student of students) {
      if (!student.canvas_user_id) {
        results.errors.push(`No Canvas ID for student: ${student.name}`);
        results.failed++;
        continue;
      }

      try {
        postGrade(courseId, discussion.canvas_assignment_id, student.canvas_user_id, grade, feedback);
        results.success++;
        Utilities.sleep(200);
      } catch (e) {
        results.errors.push(`Failed for ${student.name}: ${e.message}`);
        results.failed++;
      }
    }
  } else {
    // Individual mode: post each student's grade from StudentReports
    const reports = getApprovedUnsentReports(discussionId);

    for (const report of reports) {
      const student = getStudent(report.student_id);
      if (!student || !student.canvas_user_id) {
        results.errors.push(`No Canvas ID for student: ${report.student_name}`);
        results.failed++;
        continue;
      }

      try {
        postGrade(
          courseId,
          discussion.canvas_assignment_id,
          student.canvas_user_id,
          report.grade,
          report.feedback
        );
        updateStudentReport(report.report_id, { sent: true });
        results.success++;
        Utilities.sleep(200);
      } catch (e) {
        results.errors.push(`Failed for ${report.student_name}: ${e.message}`);
        results.failed++;
      }
    }
  }

  Logger.log(`Canvas grade posting: ${results.success} success, ${results.failed} failed`);
  return results;
}

// ============================================================================
// COURSE DATA FETCH
// ============================================================================

/**
 * Fetch comprehensive course data from Canvas.
 * Populates Students sheet and returns assignment list for teacher reference.
 *
 * @param {string} courseId
 * @param {string} classPeriod - Class period to assign to synced students
 * @returns {Object} {studentCount, assignments}
 */
function fetchCanvasCourseData(courseId, classPeriod) {
  // Fetch course info
  const course = canvasRequest(`/courses/${courseId}`);
  Logger.log(`Fetching data for course: ${course.name}`);

  // Sync students
  const studentCount = syncCanvasStudents(courseId, classPeriod);

  // Fetch assignments
  const assignments = canvasRequestPaginated(`/courses/${courseId}/assignments?order_by=due_at`);

  const assignmentList = assignments.map(a => ({
    id: a.id,
    name: a.name,
    due_at: a.due_at || 'No due date',
    points_possible: a.points_possible
  }));

  Logger.log(`Fetched ${assignmentList.length} assignments from course ${courseId}`);

  return {
    courseName: course.name,
    studentCount: studentCount,
    assignments: assignmentList
  };
}
