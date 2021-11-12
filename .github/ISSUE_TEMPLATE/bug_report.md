---
name: Bug report
about: Create a report to help us improve
title: ''
labels: ''
assignees: ''

---

**Check List**
- [ ] The App appears in the system tray
- [ ] Running `microsoft-edge:https://google.com` redirects successfully

**Describe the bug**\
A clear and concise description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Go to '...'
2. Click on '....'
3. Scroll down to '....'
4. See error

**File Upload**

/!\ Leave edge open and run `Get-WmiObject Win32_Process -Filter "name = 'msedge.exe'" | Select-Object CommandLine | Format-Table -Wrap -AutoSize | Out-File ./edge.txt` and attach the generated file

**Desktop (please complete the following information):**
 - Windows Version: [e.g. 11]
 - Windows Build: [e.g. 22494]

**Additional context**\
Add any other context about the problem here.
