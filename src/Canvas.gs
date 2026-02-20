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
 * Get sections for a Canvas course
 * @param {string} courseId
 * @returns {Object[]} Array of {id, name}
 */
function getCanvasSections(courseId) {
  const sections = canvasRequestPaginated(`/courses/${courseId}/sections`);
  return sections.map(s => ({
    id: s.id,
    name: s.name
  }));
}

/**
 * Get students enrolled in a course, with enrollment data to determine section
 * @param {string} courseId
 * @returns {Object[]} Array of student objects with sectionId
 */
function getCanvasStudents(courseId) {
  const students = canvasRequestPaginated(
    `/courses/${courseId}/users?enrollment_type[]=student&include[]=email&include[]=enrollments`
  );

  return students.map(s => {
    // Find the student enrollment for this course to get section ID
    const enrollment = (s.enrollments || []).find(
      e => e.type === 'StudentEnrollment' && e.course_id === Number(courseId)
    );
    return {
      canvasUserId: s.id,
      name: s.name,
      sortableName: s.sortable_name,
      email: s.email || '',
      loginId: s.login_id || '',
      sectionId: enrollment ? enrollment.course_section_id : null
    };
  });
}

/**
 * Sync Canvas students to local Students sheet.
 * If sectionMap is provided, auto-assigns section names from Canvas section data.
 *
 * @param {string} courseId
 * @param {string} fallbackSection - Section name to use if section can't be determined
 * @param {Object} sectionMap - Optional map of Canvas section IDs to section names
 * @param {string} courseName - Optional course name to tag students with (multi-course mode)
 * @returns {number} Number of students synced
 */
function syncCanvasStudents(courseId, fallbackSection, sectionMap = null, courseName = null) {
  const canvasStudents = getCanvasStudents(courseId);

  for (const student of canvasStudents) {
    // Determine section: use Canvas section data if available, otherwise fallback
    let section = fallbackSection;
    if (sectionMap && student.sectionId && sectionMap[student.sectionId]) {
      section = sectionMap[student.sectionId];
    }

    const studentData = {
      name: student.name,
      email: student.email,
      section: section,
      canvas_user_id: student.canvasUserId.toString()
    };
    if (courseName) {
      studentData.course = courseName;
    }

    upsertStudent(studentData);
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
// CANVAS ITEM TYPE RESOLUTION
// ============================================================================

/**
 * Get the effective Canvas item type for a discussion.
 * 3-tier fallback: per-discussion → per-course → global setting.
 *
 * @param {Object} discussion - Discussion row object
 * @returns {string} 'assignment' or 'discussion'
 */
function getCanvasItemType(discussion) {
  return getCanvasItemTypeForDiscussion(discussion);
}

/**
 * Resolve a Canvas item ID to an assignment ID for grade posting.
 * If itemType is 'discussion', fetches the discussion topic to get its assignment_id.
 * If itemType is 'assignment', returns itemId as-is.
 *
 * @param {string} courseId
 * @param {string} itemId - The ID entered by the teacher
 * @param {string} itemType - 'assignment' or 'discussion'
 * @returns {string} The assignment ID for grade posting
 */
function resolveCanvasAssignmentId(courseId, itemId, itemType) {
  if (itemType !== 'discussion') {
    return itemId;
  }

  const topic = canvasRequest(`/courses/${courseId}/discussion_topics/${itemId}`);

  if (!topic.assignment_id) {
    throw new Error(
      `Discussion topic ${itemId} is not a graded discussion (no linked assignment). ` +
      `Only graded discussion topics can receive grades.`
    );
  }

  Logger.log(`Resolved discussion topic ${itemId} to assignment ${topic.assignment_id}`);
  return topic.assignment_id.toString();
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
    throw new Error('No Canvas item ID set. Enter the Canvas assignment or discussion topic ID in the canvas_assignment_id column on the Discussions sheet.');
  }
  const canvasItemId = discussion.canvas_assignment_id;

  const courseId = getCanvasCourseIdForDiscussion(discussion);

  // Resolve item ID: if type is 'discussion', look up the linked assignment_id
  const itemType = getCanvasItemType(discussion);
  const assignmentId = resolveCanvasAssignmentId(courseId, canvasItemId, itemType);

  const results = { success: 0, failed: 0, errors: [] };

  if (isGroupMode()) {
    // Group mode: post discussion.grade + discussion.group_feedback to all students
    const grade = discussion.grade;
    const feedback = discussion.group_feedback || '';

    if (!grade) {
      throw new Error('No grade set on discussion row');
    }

    const students = discussion.section
      ? getStudentsBySectionAndCourse(discussion.section, discussion.course || null)
      : getAllRows(CONFIG.SHEETS.STUDENTS);

    for (const student of students) {
      if (!student.canvas_user_id) {
        results.errors.push(`No Canvas ID for student: ${student.name}`);
        results.failed++;
        continue;
      }

      try {
        postGrade(courseId, assignmentId, student.canvas_user_id, grade, feedback);
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
          assignmentId,
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
 * Fetches sections, auto-assigns students to their Canvas sections,
 * and returns assignment + section info for teacher reference.
 *
 * @param {string} courseId
 * @param {string} courseName - Optional course name to tag students with (multi-course mode)
 * @returns {Object} {courseName, sections, studentCount, assignments, discussionTopics}
 */
function fetchCanvasCourseData(courseId, courseName = null) {
  // Fetch course info
  const course = canvasRequest(`/courses/${courseId}`);
  Logger.log(`Fetching data for course: ${course.name}`);

  // Fetch sections and build ID → name map
  const sections = getCanvasSections(courseId);
  const sectionMap = {};
  for (const section of sections) {
    sectionMap[section.id] = section.name;
  }
  Logger.log(`Found ${sections.length} sections: ${sections.map(s => s.name).join(', ')}`);

  // Sync students with auto-section assignment
  const fallbackSection = sections.length === 1 ? sections[0].name : course.name;
  const studentCount = syncCanvasStudents(courseId, fallbackSection, sectionMap, courseName);

  // Fetch assignments
  const assignments = canvasRequestPaginated(`/courses/${courseId}/assignments?order_by=due_at`);

  const assignmentList = assignments.map(a => ({
    id: a.id,
    name: a.name,
    due_at: a.due_at || 'No due date',
    points_possible: a.points_possible
  }));

  // Fetch graded discussion topics
  const topics = canvasRequestPaginated(`/courses/${courseId}/discussion_topics?order_by=recent_activity`);

  const discussionTopicList = topics
    .filter(t => t.assignment_id)
    .map(t => ({
      id: t.id,
      name: t.title,
      assignment_id: t.assignment_id,
      due_at: t.due_at || 'No due date',
      points_possible: t.points_possible || 'N/A'
    }));

  Logger.log(`Fetched ${assignmentList.length} assignments and ${discussionTopicList.length} graded discussion topics`);

  return {
    courseName: course.name,
    sections: sections,
    studentCount: studentCount,
    assignments: assignmentList,
    discussionTopics: discussionTopicList
  };
}
