import os
import re
import subprocess

def compile_app():
    with open('Index.html', 'r', encoding='utf-8') as f:
        content = f.read()

    def replace_include(match):
        filename = match.group(1).strip()
        if not filename.endswith('.html'):
            filename += '.html'
        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as inc_f:
                return inc_f.read()
        return f'<!-- Include {filename} not found -->'

    compiled = re.sub(r'<\?!=?\s*include\(["\']([^"\']+)["\']\);?\s*\?>', replace_include, content)

    # In mock/offline mode, provide a mock google.script.run so UI scripts run smoothly
    mock_script = """
<script>
window.google = window.google || {};
window.google.script = window.google.script || {};
window.google.script.run = {
  withSuccessHandler: function(fn) {
    this._success = fn;
    return this;
  },
  withFailureHandler: function(fn) {
    this._failure = fn;
    return this;
  },
  getInitialAppState: function() {
    var self = this;
    setTimeout(function() {
      if (self._success) self._success({
        currentUser: { emp_id: 'HD1001', emp_name: 'Admin User', app_role: 'Admin', team: 'Management' },
        activeShift: null,
        clientMaster: [
          { client_name: 'iMerit', project_name: 'AMBALL' },
          { client_name: 'Hurix', project_name: 'Content QA' }
        ]
      });
    }, 10);
  },
  getAdminDashboardData: function(filters) {
    var self = this;
    setTimeout(function() {
      if (self._success) self._success({
        success: true,
        summary: {
          totalEmployees: 25,
          clockedInCount: 18,
          leaveCount: 2,
          totalLeaveHours: 16,
          totalShiftHours: 144,
          totalLoggedHours: 130,
          productionHours: 95,
          supportHours: 12,
          totalAvailableProductionHours: 144,
          productionUtilizationPct: 65.97,
          meetingHours: 10,
          trainingHours: 8,
          downtimeIdleHours: 5
        },
        employeeStatusList: [
          { emp_id: 'HD1002', emp_name: 'John Doe', team: 'Annotation', shift_type: 'General', shift_status: 'OPEN', attendance_status: 'Present', net_shift_hours: 8, leave_hours: 0, total_task_hours: 7.5, tasks: [] }
        ],
        summaryTable: { rows: [], total: {} }
      });
    }, 10);
  }
};
</script>
"""
    compiled = compiled.replace('</head>', mock_script + '\n</head>')

    with open('full_compiled.html', 'w', encoding='utf-8') as out_f:
        out_f.write(compiled)
    print("Compiled full_compiled.html successfully.")

compile_app()

