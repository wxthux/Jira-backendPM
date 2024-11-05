require('dotenv').config();
const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const fetch = require('node-fetch');
const xlsx = require('xlsx');

const app = express();
const port = 3000;

app.use(cors());
app.use(express.json());
app.use(helmet());

app.use(helmet.contentSecurityPolicy({
  directives: {
    defaultSrc: ["'self'"],
    imgSrc: ["'self'", "data:"],
  }
}));

// Default route for root URL
app.get('/', (req, res) => {
  res.send('Welcome to the JIRA Worklog Report API! Use /api/generate-report to generate reports.');
});

// Define JIRA API details
const jiraUrl = 'https://projectmanagementzone.atlassian.net/rest/api/2/search?jql=timespent > 0&fields=worklog,summary,assignee,project,key&expand=worklog';
const username = process.env.JIRA_USERNAME;;
const apiToken = process.env.JIRA_TOKEN;
const authHeader = `Basic ${Buffer.from(`${username}:${apiToken}`).toString('base64')}`;

app.post('/api/generate-report', async (req, res) => {
  const { startDate, endDate } = req.body;

  try {
    const response = await fetch(jiraUrl, {
      method: 'GET',
      headers: {
        'Authorization': authHeader,
        'Accept': 'application/json'
      }
    });

    if (!response.ok) {
      return res.status(response.status).json({ error: response.statusText });
    }

    const results = await response.json();
    const taskData = [];

    results.issues.forEach(issue => {
      const title = issue.fields.summary || "No title available";
      const assigneeEmail = issue.fields.assignee ? issue.fields.assignee.displayName : "Unassigned";
      const projectName = issue.fields.project ? issue.fields.project.name : "No project name";
      const projectId = issue.fields.project ? issue.fields.project.key : "No project ID";
      const issueId = issue.key || "No issue ID";
      issue.fields.worklog.worklogs.forEach(log => {
        const logDate = new Date(log.started);
        if (logDate >= new Date(startDate) && logDate <= new Date(endDate)) {
          
          const updatedBy = log.updateAuthor ? log.updateAuthor.displayName  : assigneeEmail;
          const comment = log.comment || "No comment";
          const timeSpent = log.timeSpentSeconds / 3600;

          taskData.push({
            Date: logDate.toISOString().split("T")[0],
            Assignee: assigneeEmail,
            ProjectName: projectName,
            ProjectID: projectId,
            IssueID: issueId,
            UpdatedBy: updatedBy, // Add updated by information
            Issue: title,
            Comment: comment,
            Hours: timeSpent.toFixed(2)
          });
        }
      });
    });

    taskData.sort((a, b) => new Date(a.Date) - new Date(b.Date));

    // Create a worksheet and style it
    const worksheet = xlsx.utils.json_to_sheet(taskData);
    
    // Set column headers style
    const headerCells = Object.keys(taskData[0]);
    headerCells.forEach((header, index) => {
      const cellAddress = xlsx.utils.encode_cell({ r: 0, c: index });
      worksheet[cellAddress].s = {
        fill: {
          fgColor: { rgb: "FFA500" } 
        },
        font: {
          bold: true
        }
      };
    });

    // Apply alternating row colors
    for (let i = 0; i < taskData.length; i++) {
      const row = i + 1; // Data starts from row 1
      const rowCells = Object.keys(taskData[i]);

      rowCells.forEach((cell, index) => {
        const cellAddress = xlsx.utils.encode_cell({ r: row, c: index });
        worksheet[cellAddress].s = {
          fill: {
            fgColor: { rgb: (i % 2 === 0) ? "FFFFFF" : "D3D3D3" } 
          }
        };
      });
    }

    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, "Worklog Report");

    const fileName = `Worklog_Report_${startDate}_to_${endDate}.xlsx`;
    xlsx.writeFile(workbook, fileName);

    res.download(fileName, (err) => {
      if (err) {
        console.error('Error downloading file:', err);
      }
    });

  } catch (error) {
    console.error('Failed to fetch or process JIRA data:', error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
