const express = require('express');
const cors = require('cors');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const FormData = require('form-data');
const { BedrockRuntimeClient, InvokeModelCommand } = require('@aws-sdk/client-bedrock-runtime');
const { google } = require('googleapis');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const { S3Client, PutObjectCommand } = require('@aws-sdk/client-s3');

require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 5000;

// CORS configuration for both React apps
app.use(cors({
  origin: '*',
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With']
}));

// Handle preflight requests
app.options('*', (req, res) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
  res.sendStatus(200);
});

app.use(express.json());

// TODO: FRIEND'S KATAL APP - Replace with actual Katal build path when available
// const katalBuildPath = path.join(__dirname, '../../ConnectorSampleDataGenerator/src/ConnectorSampleDataGeneratorWebsite/build');
// app.use('/katal', express.static(katalBuildPath));

// Serve original React build files
// const clientBuildPath = path.join(__dirname, '../client/build');
// if (fs.existsSync(clientBuildPath)) {
//   app.use('/client', express.static(clientBuildPath));
// }

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'Server running', timestamp: new Date().toISOString() });
});

// TODO: FRIEND'S KATAL ENDPOINTS - Add these when Katal app is ready
// app.get('/api/katal-data', (req, res) => {
//   // Endpoint for Katal app data
//   res.json({ message: 'Katal data endpoint' });
// });

// app.post('/api/katal-submit', (req, res) => {
//   // Endpoint for Katal form submissions
//   res.json({ success: true, message: 'Data received from Katal app' });
// });

const file_type_names = [
  "PDF", "DOCX", "PPTX", "MARKDOWN", "HTML", "TXT", "RST", 
  "PNG", "JPEG", "RTF", "XLSX", "XLS", "CSV", "TSV", "SVG"
];

function generateRelevantFile(fileType, fileName, content) {
  const filePath = path.join(__dirname, 'temp', `${fileName.replace(/[^a-zA-Z0-9]/g, '_')}_${Date.now()}.${fileType.toLowerCase()}`);
  
  // Ensure temp directory exists
  if (!fs.existsSync(path.join(__dirname, 'temp'))) {
    fs.mkdirSync(path.join(__dirname, 'temp'));
  }
  
  // Convert content to string if it's an array
  const stringContent = Array.isArray(content) ? content.join('\n') : String(content);
  
  let fileContent;
  switch (fileType) {
    case 'TXT':
    case 'MARKDOWN':
    case 'HTML':
    case 'RST':
      fileContent = stringContent;
      break;
    case 'CSV':
      fileContent = `Title,Description,Content\n"${fileName}","Generated file","${stringContent.replace(/"/g, '""')}"`;
      break;
    case 'TSV':
      fileContent = `Title\tDescription\tContent\n${fileName}\tGenerated file\t${stringContent}`;
      break;
    case 'SVG':
      fileContent = `<svg width="200" height="100" xmlns="http://www.w3.org/2000/svg"><rect width="200" height="100" fill="#4f46e5"/><text x="10" y="30" fill="white" font-size="12">${fileName}</text></svg>`;
      break;
    case 'PNG':
    case 'JPEG':
      const imageData = fileType === 'PNG' ? 
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==' :
        '/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwA/wA==';
      fs.writeFileSync(filePath, Buffer.from(imageData, 'base64'));
      return filePath;
    default:
      fileContent = stringContent;
  }
  
  fs.writeFileSync(filePath, fileContent);
  return filePath;
}

function generateRandomFile(fileType, issueTitle) {
  const fileName = `${issueTitle.replace(/[^a-zA-Z0-9]/g, '_')}_${Date.now()}`;
  const filePath = path.join(__dirname, 'temp', `${fileName}.${fileType.toLowerCase()}`);
  
  // Ensure temp directory exists
  if (!fs.existsSync(path.join(__dirname, 'temp'))) {
    fs.mkdirSync(path.join(__dirname, 'temp'));
  }
  
  let content;
  switch (fileType) {
    case 'TXT':
    case 'MARKDOWN':
    case 'HTML':
    case 'RST':
      content = `# ${issueTitle}\n\nThis is a sample ${fileType} file for the issue.\nGenerated on: ${new Date().toISOString()}\n\nContent related to: ${issueTitle}`;
      break;
    case 'CSV':
      content = `Title,Description,Status\n"${issueTitle}","Sample data","In Progress"\n"Related Item","Additional info","Done"`;
      break;
    case 'TSV':
      content = `Title\tDescription\tStatus\n${issueTitle}\tSample data\tIn Progress\nRelated Item\tAdditional info\tDone`;
      break;
    case 'SVG':
      content = `<svg width="200" height="100" xmlns="http://www.w3.org/2000/svg"><rect width="200" height="100" fill="#4f46e5"/><text x="10" y="30" fill="white" font-size="12">${issueTitle}</text></svg>`;
      break;
    case 'PNG':
    case 'JPEG':
      // Create a simple base64 encoded 1x1 pixel image
      const imageData = fileType === 'PNG' ? 
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==' :
        '/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAAAAX/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwA/wA==';
      fs.writeFileSync(filePath, Buffer.from(imageData, 'base64'));
      return filePath;
    default:
      content = `Sample ${fileType} file for: ${issueTitle}\nGenerated: ${new Date().toISOString()}`;
  }
  
  fs.writeFileSync(filePath, content);
  return filePath;
}



app.post('/api/bedrock', async (req, res) => {
  try {
    const { prompt } = req.body;
    const input = {
      modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
      contentType: 'application/json',
      accept: 'application/json',
      body: JSON.stringify({
        anthropic_version: 'bedrock-2023-05-31',
        max_tokens: 1000,
        messages: [{ role: 'user', content: prompt }]
      })
    };
    const command = new InvokeModelCommand(input);
    const response = await client.send(command);
    const responseBody = JSON.parse(new TextDecoder().decode(response.body));
    res.json({
      success: true,
      response: responseBody.content[0].text
    });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

app.post('/api/s3/create-files', async (req, res) => {
  try {
    const { prompt, bucketName, fileCount = 10, awsCredentials } = req.body;
    if (!bucketName || !awsCredentials) {
      return res.status(400).json({
        success: false,
        error: 'Bucket name and AWS credentials are required'
      });
    }
    // Create S3 client with user-provided credentials
    const userS3Client = new S3Client({
      region: awsCredentials.region || 'us-east-2',
      credentials: {
        accessKeyId: awsCredentials.accessKeyId,
        secretAccessKey: awsCredentials.secretAccessKey,
        ...(awsCredentials.sessionToken && { sessionToken: awsCredentials.sessionToken })
      }
    });
    const file_type_names = ["TXT", "MARKDOWN", "HTML", "CSV", "RTF"];
    const teamFolders = ['hr', 'legal', 'policies', 'documentation', 'finance'];
    const createdFiles = [];
    const failedFiles = [];
    // Generate project name from prompt using AI
    const namePrompt = `Extract a company name or project name from this prompt: "${prompt}"
Return ONLY a short, clean name (2-4 words max) suitable for a folder name. If no specific company/project name is found, create a relevant project name based on the context.
Examples:
    - "Create marketing materials for TechCorp" → "TechCorp"
    - "Develop HR policies for startup" → "HR-Policies-Project"
Return only the name:`;
    let projectName;
    try {
      const nameInput = {
        modelId: 'anthropic.claude-3-5-sonnet-20240620-v1:0',
        contentType: 'application/json',
        accept: 'application/json',
        body: JSON.stringify({
          anthropic_version: 'bedrock-2023-05-31',
          max_tokens: 100,
          messages: [{ role: 'user', content: namePrompt }]
        })
      };
      const nameCommand = new InvokeModelCommand(nameInput);
      const nameResponse = await client.send(nameCommand);
      const nameBody = JSON.parse(new TextDecoder().decode(nameResponse.body));
      projectName = nameBody.content[0].text.trim().replace(/[^a-zA-Z0-9\s-]/g, '').replace(/\s+/g, '-').toLowerCase().substring(0, 25);
    } catch (error) {
      projectName = prompt.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '-').toLowerCase().substring(0, 20);
    }
    for (let i = 0; i < fileCount; i++) {
      try {
        const randomTeam = teamFolders[Math.floor(Math.random() * teamFolders.length)];
        const randomFileType = file_type_names[Math.floor(Math.random() * file_type_names.length)];
        const filePrompt = `Generate professional ${randomFileType} content for ${randomTeam} team based on this issue: "${prompt}".
Requirements:
- Write exactly 100-200 words
- Make it relevant and specific to the issue
- Use professional business language
- Include actionable information related to resolving the issue
- Format appropriately for ${randomFileType}
Return only the content, no explanations:`;
        const input = {
          modelId: 'anthropic.claude-3-5-sonnet-20240620-v1:0',
          contentType: 'application/json',
          accept: 'application/json',
          body: JSON.stringify({
            anthropic_version: 'bedrock-2023-05-31',
            max_tokens: 500,
            messages: [{ role: 'user', content: filePrompt }]
          })
        };
        const command = new InvokeModelCommand(input);
        const aiResponse = await client.send(command);
        const responseBody = JSON.parse(new TextDecoder().decode(aiResponse.body));
        const content = responseBody.content[0].text.trim();
        const extensions = { TXT: '.txt', MARKDOWN: '.md', HTML: '.html', CSV: '.csv', RTF: '.rtf' };
        const fileName = `${randomTeam}_${i + 1}${extensions[randomFileType]}`;
        // S3 key with nested structure: s3-q-connector/project-name/team/filename
        const s3Key = `s3-q-connector/${projectName}/${randomTeam}/${fileName}`;
        const uploadCommand = new PutObjectCommand({
          Bucket: bucketName,
          Key: s3Key,
          Body: content,
          ContentType: 'text/plain'
        });
        await userS3Client.send(uploadCommand);
        createdFiles.push({
          name: fileName,
          team: randomTeam,
          fileType: randomFileType,
          size: content.length,
          s3Key: s3Key,
          url: `https://${bucketName}.s3.${awsCredentials.region || 'us-east-1'}.amazonaws.com/${s3Key}`
        });
      } catch (fileError) {
        console.error(`Failed to create file ${i + 1}:`, fileError.message);
        failedFiles.push({
          index: i + 1,
          error: fileError.message
        });
      }
    }
    res.json({
      success: createdFiles.length > 0,
      filesCreated: createdFiles.length,
      files: createdFiles,
      failedFiles: failedFiles,
      projectFolder: `s3-q-connector/${projectName}`,
      bucketName: bucketName,
      message: `Created ${createdFiles.length} files in S3 nested structure`
    });
  } catch (error) {
    console.error('S3 error:', error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

// jira
app.post('/api/create-issues', async (req, res) => {
  try {
    const { prompt, jiraConfig, issueCount = 3, projectName, projectKey, securityLevel = 'Mixed' } = req.body;
    const auth = Buffer.from(`${jiraConfig.email}:${jiraConfig.apiToken}`).toString('base64');
    
    let finalProjectKey = projectKey || jiraConfig.projectKey;
    
    // Create project if projectName and projectKey are provided
    if (projectName && projectKey) {
      try {
        // Get current user account ID
        const userResponse = await axios.get(
          `${jiraConfig.baseUrl}/rest/api/3/myself`,
          { headers: { Authorization: `Basic ${auth}` } }
        );
        
        const newProject = {
          key: projectKey,
          name: projectName,
          projectTypeKey: 'software',
          description: `Project created for: ${prompt}`,
          leadAccountId: userResponse.data.accountId
        };
        
        await axios.post(
          `${jiraConfig.baseUrl}/rest/api/3/project`,
          newProject,
          { headers: { Authorization: `Basic ${auth}`, 'Content-Type': 'application/json' } }
        );
        
        finalProjectKey = projectKey;
        console.log(`Created project: ${projectName} (${projectKey})`);
      } catch (projectError) {
        console.log('Project might already exist or creation failed:', projectError.response?.data);
        finalProjectKey = projectKey; // Use provided key anyway
      }
    }
    
    // Get available issue types for the project
    const projectResponse = await axios.get(
      `${jiraConfig.baseUrl}/rest/api/3/project/${finalProjectKey}`,
      { headers: { Authorization: `Basic ${auth}` } }
    );
    const availableIssueTypes = projectResponse.data.issueTypes;
    const defaultIssueType = availableIssueTypes[0]; // Use first available issue type
    
    // Generate issues with AI
    const aiPrompt = `You are a Jira issue generator. Create exactly ${issueCount} distinct, actionable Jira issues for: "${prompt}"

Requirements:
- Each issue must have a clear, specific title (max 100 chars)
- Each description must be detailed and actionable (2-3 sentences)
- Make issues diverse and cover different aspects
- Focus on practical, implementable tasks

Return ONLY this JSON array format:
[{"title": "Issue title here", "description": "Detailed description here"}, ...]
Create issues title unique .
No explanations, no markdown, just the JSON array:`;
    const input = {
      modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
      contentType: 'application/json',
      accept: 'application/json',
      body: JSON.stringify({
        anthropic_version: 'bedrock-2023-05-31',
        max_tokens: 2000,
        messages: [{ role: 'user', content: aiPrompt }]
      })
    };
    const command = new InvokeModelCommand(input);
    const response = await client.send(command);
    const responseBody = JSON.parse(new TextDecoder().decode(response.body));
    
    const issuesText = responseBody.content[0].text;
    console.log('AI Response:', issuesText);
    
    let issues;
    try {
      const jsonMatch = issuesText.match(/\[.*\]/s);
      if (!jsonMatch) {
        throw new Error('No JSON array found in AI response');
      }
      issues = JSON.parse(jsonMatch[0]);
    } catch (parseError) {
      console.error('Failed to parse AI response:', parseError);
      // Fallback: create simple issues from the prompt
      issues = Array.from({ length: issueCount }, (_, i) => ({
        title: `${prompt} - Issue ${i + 1}`,
        description: `Generated issue ${i + 1} based on: ${prompt}`
      }));
    }
    
    // Create issues in Jira
    const createdIssues = [];
    
    for (const issue of issues) {
      try {
        // Add security level as label
        let securityLabels = [];
        if (securityLevel === 'Mixed') {
          securityLabels = Math.random() > 0.5 ? ['confidential'] : ['open'];
        } else if (securityLevel === 'Confidential') {
          securityLabels = ['confidential'];
        } else if (securityLevel === 'Open') {
          securityLabels = ['open'];
        }
        
        const jiraIssue = {
          fields: {
            project: { key: finalProjectKey },
            summary: issue.title,
            description: {
              type: "doc",
              version: 1,
              content: [{
                type: "paragraph",
                content: [{
                  type: "text",
                  text: issue.description
                }]
              }]
            },
            issuetype: { id: defaultIssueType.id },
            labels: securityLabels
          }
        };
        
        const jiraResponse = await axios.post(
          `${jiraConfig.baseUrl}/rest/api/3/issue`,
          jiraIssue,
          { headers: { Authorization: `Basic ${auth}`, 'Content-Type': 'application/json' } }
        );
        
        const issueKey = jiraResponse.data.key;
        console.log(`Created issue: ${issueKey} - ${issue.title}`);
        
        // Wait for issue to be fully created before adding comments/worklogs
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        // Add random attachments to issue (1-2 attachments)
        const attachmentCount = Math.floor(Math.random() * 2) + 1;
        for (let i = 0; i < attachmentCount; i++) {
          const randomFileType = file_type_names[Math.floor(Math.random() * file_type_names.length)];
          const filePath = generateRandomFile(randomFileType, issue.title);
          
          try {
            const formData = new FormData();
            formData.append('file', fs.createReadStream(filePath));
            
            await axios.post(
              `${jiraConfig.baseUrl}/rest/api/3/issue/${issueKey}/attachments`,
              formData,
              {
                headers: {
                  ...formData.getHeaders(),
                  'Authorization': `Basic ${auth}`,
                  'X-Atlassian-Token': 'no-check'
                }
              }
            );
            
            // Clean up temp file
            fs.unlinkSync(filePath);
          } catch (attachError) {
            console.error('Failed to add attachment:', attachError.response?.data);
            if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
          }
        }
        
        // Add random comments (1-4 comments per issue)
        const commentCount = Math.floor(Math.random() * 4) + 1;
        for (let i = 0; i < commentCount; i++) {
          const hasAttachment = Math.random() > 0.6; // 40% chance of attachment
          const commentPrompt = hasAttachment ? 
            `Generate a realistic Jira comment mentioning an attached ${file_type_names[Math.floor(Math.random() * file_type_names.length)].toLowerCase()} file for issue: "${issue.title}". Make it professional. Return only the comment text.` :
            `Generate a realistic Jira comment for issue: "${issue.title}". Make it professional, specific, and related to the issue. Return only the comment text, no quotes or formatting.`;
          
          try {
            const commentInput = {
              modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
              contentType: 'application/json',
              accept: 'application/json',
              body: JSON.stringify({
                anthropic_version: 'bedrock-2023-05-31',
                max_tokens: 200,
                messages: [{ role: 'user', content: commentPrompt }]
              })
            };
            const commentCommand = new InvokeModelCommand(commentInput);
            const commentResponse = await client.send(commentCommand);
            const commentBody = JSON.parse(new TextDecoder().decode(commentResponse.body));
            const commentText = commentBody.content[0].text.trim();
            
            const commentResponse2 = await axios.post(
              `${jiraConfig.baseUrl}/rest/api/3/issue/${issueKey}/comment`,
              {
                body: {
                  type: "doc",
                  version: 1,
                  content: [{
                    type: "paragraph",
                    content: [{
                      type: "text",
                      text: commentText
                    }]
                  }]
                }
              },
              { headers: { Authorization: `Basic ${auth}`, 'Content-Type': 'application/json' } }
            );
            
            console.log(`Added comment to ${issueKey}:`, commentText.substring(0, 50) + '...');
            
            // Add attachment to comment if needed
            if (hasAttachment) {
              const randomFileType = file_type_names[Math.floor(Math.random() * file_type_names.length)];
              const filePath = generateRandomFile(randomFileType, `comment_${issue.title}`);
              
              try {
                const formData = new FormData();
                formData.append('file', fs.createReadStream(filePath));
                
                await axios.post(
                  `${jiraConfig.baseUrl}/rest/api/3/issue/${issueKey}/attachments`,
                  formData,
                  {
                    headers: {
                      ...formData.getHeaders(),
                      'Authorization': `Basic ${auth}`,
                      'X-Atlassian-Token': 'no-check'
                    }
                  }
                );
                
                fs.unlinkSync(filePath);
              } catch (attachError) {
                console.error('Failed to add comment attachment:', attachError.response?.data);
                if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
              }
            }
          } catch (commentError) {
            console.error('Failed to add comment:', commentError.response?.data || commentError.message);
          }
          
          // Small delay between comments
          await new Promise(resolve => setTimeout(resolve, 500));
        }
        
        // Add random worklogs (1-3 worklogs per issue)
        const worklogCount = Math.floor(Math.random() * 3) + 1;
        for (let i = 0; i < worklogCount; i++) {
          const timeSpent = Math.floor(Math.random() * 6) + 1; // 1-6 hours
          const worklogPrompt = `Generate a brief work description for ${timeSpent} hours of work on: "${issue.title}". Make it specific and professional. Return only the description, no quotes.`;
          
          try {
            const worklogInput = {
              modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
              contentType: 'application/json',
              accept: 'application/json',
              body: JSON.stringify({
                anthropic_version: 'bedrock-2023-05-31',
                max_tokens: 150,
                messages: [{ role: 'user', content: worklogPrompt }]
              })
            };
            const worklogCommand = new InvokeModelCommand(worklogInput);
            const worklogResponse = await client.send(worklogCommand);
            const worklogBody = JSON.parse(new TextDecoder().decode(worklogResponse.body));
            const worklogDescription = worklogBody.content[0].text.trim();
            
            const startDate = new Date();
            startDate.setDate(startDate.getDate() - Math.floor(Math.random() * 7)); // Random date within last week
            
            await axios.post(
              `${jiraConfig.baseUrl}/rest/api/3/issue/${issueKey}/worklog`,
              {
                timeSpent: `${timeSpent}h`,
                started: startDate.toISOString().replace('Z', '+0000'),
                comment: {
                  type: "doc",
                  version: 1,
                  content: [{
                    type: "paragraph",
                    content: [{
                      type: "text",
                      text: worklogDescription
                    }]
                  }]
                }
              },
              { headers: { Authorization: `Basic ${auth}`, 'Content-Type': 'application/json' } }
            );
            
            console.log(`Added ${timeSpent}h worklog to ${issueKey}:`, worklogDescription.substring(0, 50) + '...');
          } catch (worklogError) {
            console.error('Failed to add worklog:', worklogError.response?.data || worklogError.message);
          }
          
          // Small delay between worklogs
          await new Promise(resolve => setTimeout(resolve, 500));
        }
        
        createdIssues.push(jiraResponse.data);
      } catch (issueError) {
        console.error('Failed to create issue:', issue.title, issueError.response?.data);
      }
    }
    
    res.json({
      success: true,
      message: `Created ${createdIssues.length} issues successfully using issue type: ${defaultIssueType.name}`,
      issues: createdIssues
    });
  } catch (error) {
    console.error('Full error:', error.response?.data || error.message);
    res.status(500).json({ success: false, error: error.response?.data || error.message });
  }
});

// OneDrive file creation with nested folders
app.post('/api/onedrive/create-files', async (req, res) => {
  try {
    const { prompt, accessToken, fileCount = 10 } = req.body;
    
    if (!accessToken) {
      return res.status(400).json({ 
        success: false, 
        error: 'Access token is required' 
      });
    }
    
    const file_type_names = ["TXT", "MARKDOWN", "HTML", "CSV", "RTF"];
    const teamFolders = ['hr', 'legal', 'policies', 'documentation', 'finance'];
    
    const createdFiles = [];
    const failedFiles = [];
    const folderCache = {};
    
    // Helper function to create folder
    const createFolder = async (folderName, parentId = null) => {
      try {
        const folderData = {
          name: folderName,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'rename'
        };
        
        const url = parentId 
          ? `https://graph.microsoft.com/v1.0/me/drive/items/${parentId}/children`
          : 'https://graph.microsoft.com/v1.0/me/drive/root/children';
        
        const response = await axios.post(url, folderData, {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        });
        
        return response.data.id;
      } catch (error) {
        if (error.response?.status === 409) {
          // Folder exists, get its ID
          const searchUrl = parentId 
            ? `https://graph.microsoft.com/v1.0/me/drive/items/${parentId}/children?$filter=name eq '${folderName}'`
            : `https://graph.microsoft.com/v1.0/me/drive/root/children?$filter=name eq '${folderName}'`;
          
          const searchResponse = await axios.get(searchUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
          });
          
          return searchResponse.data.value[0]?.id;
        }
        throw error;
      }
    };
    
    // Generate project name from prompt using AI
    const namePrompt = `Extract a company name or project name from this prompt: "${prompt}"
    
Return ONLY a short, clean name (2-4 words max) suitable for a folder name. If no specific company/project name is found, create a relevant project name based on the context.
    
Examples:
    - "Create marketing materials for TechCorp" → "TechCorp"
    - "Develop HR policies for startup" → "HR-Policies-Project"
    
Return only the name:`;
    
    let projectName;
    try {
      const nameInput = {
        modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
        contentType: 'application/json',
        accept: 'application/json',
        body: JSON.stringify({
          anthropic_version: 'bedrock-2023-05-31',
          max_tokens: 100,
          messages: [{ role: 'user', content: namePrompt }]
        })
      };
      
      const nameCommand = new InvokeModelCommand(nameInput);
      const nameResponse = await client.send(nameCommand);
      const nameBody = JSON.parse(new TextDecoder().decode(nameResponse.body));
      projectName = nameBody.content[0].text.trim().replace(/[^a-zA-Z0-9\s-]/g, '').replace(/\s+/g, '-').toLowerCase().substring(0, 25);
    } catch (error) {
      projectName = prompt.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '-').toLowerCase().substring(0, 20);
    }
    
    const mainFolderId = await createFolder('onedrive-q-connector');
    const projectFolderId = await createFolder(projectName, mainFolderId);
    
    for (let i = 0; i < fileCount; i++) {
      try {
        const randomTeam = teamFolders[Math.floor(Math.random() * teamFolders.length)];
        const randomFileType = file_type_names[Math.floor(Math.random() * file_type_names.length)];
        
        // Create team folder if not exists
        if (!folderCache[randomTeam]) {
          folderCache[randomTeam] = await createFolder(randomTeam, projectFolderId);
        }
        
        const filePrompt = `Generate professional ${randomFileType} content for ${randomTeam} team related to: "${prompt}". 
        
Requirements:
- Write exactly 100-200 words
- Make it relevant and specific to the prompt
- Use professional business language
- Include actionable information
- Format appropriately for ${randomFileType}

Return only the content, no explanations:`;
        
        const input = {
          modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
          contentType: 'application/json',
          accept: 'application/json',
          body: JSON.stringify({
            anthropic_version: 'bedrock-2023-05-31',
            max_tokens: 500,
            messages: [{ role: 'user', content: filePrompt }]
          })
        };
        
        const command = new InvokeModelCommand(input);
        const aiResponse = await client.send(command);
        const responseBody = JSON.parse(new TextDecoder().decode(aiResponse.body));
        const content = responseBody.content[0].text.trim();
        
        const extensions = { TXT: '.txt', MARKDOWN: '.md', HTML: '.html', CSV: '.csv', RTF: '.rtf' };
        const fileName = `${randomTeam}_${i + 1}${extensions[randomFileType]}`;
        
        // Upload file to team folder
        const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${folderCache[randomTeam]}:/${fileName}:/content`;
        
        const uploadResponse = await axios.put(uploadUrl, content, {
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'text/plain'
          }
        });
        
        if (uploadResponse.status === 200 || uploadResponse.status === 201) {
          createdFiles.push({
            name: fileName,
            team: randomTeam,
            fileType: randomFileType,
            size: content.length,
            id: uploadResponse.data.id,
            webUrl: uploadResponse.data.webUrl
          });
        }
        
      } catch (fileError) {
        console.error(`Failed to create file ${i + 1}:`, fileError.response?.data || fileError.message);
        failedFiles.push({
          index: i + 1,
          error: fileError.response?.data?.error?.message || fileError.message
        });
      }
    }
    
    res.json({
      success: createdFiles.length > 0,
      filesCreated: createdFiles.length,
      files: createdFiles,
      failedFiles: failedFiles,
      projectFolder: `onedrive-q-connector/${projectName}`,
      message: `Created ${createdFiles.length} files in nested structure`
    });
  } catch (error) {
    console.error('OneDrive error:', error.response?.data || error.message);
    
    if (error.response?.data?.error?.code === 'InvalidAuthenticationToken') {
      res.status(401).json({ 
        success: false, 
        error: 'Access token expired or invalid. Please re-authenticate with OneDrive.',
        code: 'TOKEN_EXPIRED'
      });
    } else if (error.message.includes('Invalid character in header')) {
      res.status(400).json({ success: false, error: 'Invalid access token format' });
    } else {
      res.status(500).json({ success: false, error: error.response?.data?.error?.message || error.message });
    }
  }
});

// AI-Generated Google Drive Project with Relevant Content
app.post('/api/upload-gdrive-files', async (req, res) => {
  try {
    const { accessToken, fileCount = 10, prompt } = req.body;
    
    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: accessToken });
    const drive = google.drive({ version: 'v3', auth });
    
    const createdItems = [];
    const folderStructure = {};
    
    // Generate project name from prompt
    const namePrompt = `Extract a clean project name from: "${prompt}"
Return ONLY the name (2-4 words max), no quotes:`;
    
    let projectName;
    try {
      const nameInput = {
        modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
        contentType: 'application/json',
        accept: 'application/json',
        body: JSON.stringify({
          anthropic_version: 'bedrock-2023-05-31',
          max_tokens: 50,
          messages: [{ role: 'user', content: namePrompt }]
        })
      };
      
      const nameCommand = new InvokeModelCommand(nameInput);
      const nameResponse = await client.send(nameCommand);
      const nameBody = JSON.parse(new TextDecoder().decode(nameResponse.body));
      projectName = nameBody.content[0].text.trim().replace(/[^a-zA-Z0-9\s-]/g, '').replace(/\s+/g, '-');
    } catch {
      projectName = prompt.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '-').substring(0, 20);
    }
    
    // Create folder helper
    const createFolder = async (folderName, parentId = null) => {
      const folder = await drive.files.create({
        requestBody: {
          name: folderName,
          mimeType: 'application/vnd.google-apps.folder',
          parents: parentId ? [parentId] : undefined
        }
      });
      return folder.data.id;
    };
    
    // Create main project structure
    const mainFolderId = await createFolder('gdrive-q-connector');
    const projectFolderId = await createFolder(projectName, mainFolderId);
    
    // Create organized subfolders
    const subfolders = {
      'Documents': await createFolder('Documents', projectFolderId),
      'Resources': await createFolder('Resources', projectFolderId),
      'Data': await createFolder('Data', projectFolderId),
      'Reports': await createFolder('Reports', projectFolderId)
    };
    
    // Create additional subfolders in each main folder
    const docSubfolder = await createFolder('Templates', subfolders['Documents']);
    const resourceSubfolder = await createFolder('Images', subfolders['Resources']);
    const dataSubfolder = await createFolder('Raw', subfolders['Data']);
    const reportSubfolder = await createFolder('Monthly', subfolders['Reports']);
    
    // Extended folder structure for file placement
    folderStructure['Documents'] = subfolders['Documents'];
    folderStructure['Documents-Templates'] = docSubfolder;
    folderStructure['Resources'] = subfolders['Resources'];
    folderStructure['Resources-Images'] = resourceSubfolder;
    folderStructure['Data'] = subfolders['Data'];
    folderStructure['Data-Raw'] = dataSubfolder;
    folderStructure['Reports'] = subfolders['Reports'];
    folderStructure['Reports-Monthly'] = reportSubfolder;
    
    // Generate and upload files
    const folderKeys = Object.keys(folderStructure);
    
    for (let i = 0; i < fileCount; i++) {
      try {
        const randomFileType = file_type_names[Math.floor(Math.random() * file_type_names.length)];
        const folderKey = folderKeys[Math.floor(Math.random() * folderKeys.length)];
        const folderId = folderStructure[folderKey];
        
        // Generate relevant file content
        const filePrompt = `Generate ${randomFileType} content for ${folderKey.replace('-', ' ')} folder related to: "${prompt}"
Write 100-150 words of relevant professional content. Return only the content:`;
        
        const fileInput = {
          modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
          contentType: 'application/json',
          accept: 'application/json',
          body: JSON.stringify({
            anthropic_version: 'bedrock-2023-05-31',
            max_tokens: 400,
            messages: [{ role: 'user', content: filePrompt }]
          })
        };
        
        const fileCommand = new InvokeModelCommand(fileInput);
        const fileResponse = await client.send(fileCommand);
        const fileBody = JSON.parse(new TextDecoder().decode(fileResponse.body));
        const fileContent = fileBody.content[0].text.trim();
        
        const fileName = `${folderKey.toLowerCase()}_${i + 1}`;
        const filePath = generateRelevantFile(randomFileType, fileName, fileContent);
        
        const uploadedFile = await drive.files.create({
          requestBody: {
            name: `${fileName}.${randomFileType.toLowerCase()}`,
            parents: [folderId]
          },
          media: {
            mimeType: 'application/octet-stream',
            body: fs.createReadStream(filePath)
          }
        });
        
        createdItems.push({
          id: uploadedFile.data.id,
          name: uploadedFile.data.name,
          folder: folderKey,
          type: randomFileType,
          path: `gdrive-q-connector/${projectName}/${folderKey.replace('-', '/')}`
        });
        
        fs.unlinkSync(filePath);
      } catch (fileError) {
        console.error('Failed to upload file:', fileError);
      }
    }
    
    res.json({
      success: true,
      message: `Created organized folder structure with ${createdItems.length} files`,
      projectPath: `gdrive-q-connector/${projectName}`,
      files: createdItems,
      folderStructure: {
        main: 'gdrive-q-connector',
        project: projectName,
        subfolders: Object.keys(folderStructure)
      }
    });
    
  } catch (error) {
    console.error('Google Drive upload error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// TODO: FRIEND'S KATAL ROUTES - Uncomment when Katal build is available
// app.get('/katal/*', (req, res) => {
//   res.sendFile(path.join(katalBuildPath, 'index.html'));
// });

// Route for original client app
app.get('/client/*', (req, res) => {
  if (fs.existsSync(clientBuildPath)) {
    res.sendFile(path.join(clientBuildPath, 'index.html'));
  } else {
    res.status(404).json({ error: 'Client build not found' });
  }
});

// Default route - serve original client for now
app.get('/', (req, res) => {
  if (fs.existsSync(clientBuildPath)) {
    res.sendFile(path.join(clientBuildPath, 'index.html'));
  } else {
    res.json({ message: 'Server running - no frontend build available' });
  }
});

// Export app for Lambda
module.exports = app;

// Start server only if not in Lambda environment
if (require.main === module) {
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
    console.log(`Server accessible at: http://0.0.0.0:${PORT}`);
    console.log(`APIs available at: http://0.0.0.0:${PORT}/api/*`);
  });
}

// SharePoint file creation with site content updates
app.post('/api/sharepoint/create-files', async (req, res) => {
  try {
    const { 
      prompt, 
      accessToken, 
      siteUrl, 
      libraryName = 'Documents',
      fileCount = 10, 
      createFolders = false,
      updateSiteContent = false 
    } = req.body;
    
    if (!accessToken || !siteUrl) {
      return res.status(400).json({ 
        success: false, 
        error: 'Access token and site URL are required' 
      });
    }
    
    const file_type_names = ["TXT", "MARKDOWN", "HTML", "CSV", "RTF"];
    const teamFolders = ['Documents', 'Resources', 'Reports'];
    const createdFiles = [];
    const failedFiles = [];
    const folderCache = {};
    
    // Get site ID first
    const urlObj = new URL(siteUrl);
    const hostname = urlObj.hostname;
    const sitePath = urlObj.pathname;
    
    // Get site information
    const siteInfoResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`,
      { headers: { 'Authorization': `Bearer ${accessToken}` } }
    );
    
    const siteId = siteInfoResponse.data.id;
    const baseUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}`;
    
    // Generate project name
    const namePrompt = `Extract a project name from: "${prompt}"\nReturn ONLY the name (2-4 words max):`;
    let projectName;
    try {
      const nameInput = {
        modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
        contentType: 'application/json',
        accept: 'application/json',
        body: JSON.stringify({
          anthropic_version: 'bedrock-2023-05-31',
          max_tokens: 50,
          messages: [{ role: 'user', content: namePrompt }]
        })
      };
      const nameCommand = new InvokeModelCommand(nameInput);
      const nameResponse = await client.send(nameCommand);
      const nameBody = JSON.parse(new TextDecoder().decode(nameResponse.body));
      projectName = nameBody.content[0].text.trim().replace(/[^a-zA-Z0-9\s-]/g, '').replace(/\s+/g, '-');
    } catch {
      projectName = prompt.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '-').substring(0, 20);
    }
    
    // Create folder helper
    const createFolder = async (folderName, parentPath = '') => {
      try {
        const folderPath = parentPath ? `${parentPath}/${folderName}` : folderName;
        const response = await axios.post(
          `${baseUrl}/drive/root${parentPath ? `:/${parentPath}:` : ''}/children`,
          {
            name: folderName,
            folder: {},
            '@microsoft.graph.conflictBehavior': 'rename'
          },
          {
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'application/json'
            }
          }
        );
        return { id: response.data.id, path: folderPath };
      } catch (error) {
        if (error.response?.status === 409) {
          // Folder exists, get its ID
          const searchResponse = await axios.get(
            `${baseUrl}/drive/root${parentPath ? `:/${parentPath}:` : ''}/children?$filter=name eq '${folderName}'`,
            { headers: { 'Authorization': `Bearer ${accessToken}` } }
          );
          const folder = searchResponse.data.value[0];
          return { id: folder?.id, path: `${parentPath}/${folderName}` };
        }
        throw error;
      }
    };
    
    // Create main project folder in specified library
    const mainFolder = await createFolder(`${libraryName}-q-connector-${projectName}`);
    
    // Create nested folder structure if enabled
    if (createFolders) {
      for (const team of teamFolders) {
        const teamFolder = await createFolder(team, mainFolder.path);
        folderCache[team] = teamFolder;
        
        // Create sub-folders (3 levels deep)
        const subFolders = ['Active', 'Archive', 'Templates'];
        for (const sub of subFolders) {
          const subFolder = await createFolder(sub, teamFolder.path);
          folderCache[`${team}-${sub}`] = subFolder;
        }
      }
    } else {
      for (const team of teamFolders) {
        const teamFolder = await createFolder(team, mainFolder.path);
        folderCache[team] = teamFolder;
      }
    }
    
    // Generate and upload files
    for (let i = 0; i < fileCount; i++) {
      try {
        const randomTeam = teamFolders[Math.floor(Math.random() * teamFolders.length)];
        const randomFileType = file_type_names[Math.floor(Math.random() * file_type_names.length)];
        
        // Choose folder (nested or flat)
        let targetFolder;
        if (createFolders) {
          const subFolders = ['Active', 'Archive', 'Templates'];
          const randomSub = subFolders[Math.floor(Math.random() * subFolders.length)];
          targetFolder = folderCache[`${randomTeam}-${randomSub}`];
        } else {
          targetFolder = folderCache[randomTeam];
        }
        
        const filePrompt = `Generate professional ${randomFileType} content for ${randomTeam} team related to: "${prompt}". Write 100-150 words. Return only the content:`;
        
        const input = {
          modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
          contentType: 'application/json',
          accept: 'application/json',
          body: JSON.stringify({
            anthropic_version: 'bedrock-2023-05-31',
            max_tokens: 400,
            messages: [{ role: 'user', content: filePrompt }]
          })
        };
        
        const command = new InvokeModelCommand(input);
        const aiResponse = await client.send(command);
        const responseBody = JSON.parse(new TextDecoder().decode(aiResponse.body));
        const content = responseBody.content[0].text.trim();
        
        const extensions = { TXT: '.txt', MARKDOWN: '.md', HTML: '.html', CSV: '.csv', RTF: '.rtf' };
        const fileName = `${randomTeam}_${i + 1}${extensions[randomFileType]}`;
        
        // Upload file
        const uploadResponse = await axios.put(
          `${baseUrl}/drive/items/${targetFolder.id}:/${fileName}:/content`,
          content,
          {
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'text/plain'
            }
          }
        );
        
        createdFiles.push({
          name: fileName,
          team: randomTeam,
          fileType: randomFileType,
          size: content.length,
          path: `${targetFolder.path}/${fileName}`,
          webUrl: uploadResponse.data.webUrl
        });
        
      } catch (fileError) {
        failedFiles.push({
          index: i + 1,
          error: fileError.response?.data?.error?.message || fileError.message
        });
      }
    }
    
    // Update site homepage and navigation if enabled
    let siteUpdates = {};
    if (updateSiteContent) {
      try {
        // Generate homepage content
        const homepagePrompt = `Create a professional SharePoint homepage overview for project: "${prompt}". Include project description, key objectives, and team structure. Write 200-300 words in HTML format:`;
        
        const homepageInput = {
          modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
          contentType: 'application/json',
          accept: 'application/json',
          body: JSON.stringify({
            anthropic_version: 'bedrock-2023-05-31',
            max_tokens: 600,
            messages: [{ role: 'user', content: homepagePrompt }]
          })
        };
        
        const homepageCommand = new InvokeModelCommand(homepageInput);
        const homepageResponse = await client.send(homepageCommand);
        const homepageBody = JSON.parse(new TextDecoder().decode(homepageResponse.body));
        const homepageContent = homepageBody.content[0].text.trim();
        
        // Update site homepage
        const pagesResponse = await axios.get(
          `${baseUrl}/pages`,
          { headers: { 'Authorization': `Bearer ${accessToken}` } }
        );
        
        const homePage = pagesResponse.data.value.find(page => page.name === 'Home.aspx');
        if (homePage) {
          await axios.patch(
            `${baseUrl}/pages/${homePage.id}`,
            {
              title: `${projectName} - Project Overview`,
              description: `Generated content for ${projectName}`
            },
            {
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
              }
            }
          );
        }
        
        siteUpdates.homepage = {
          updated: true,
          title: `${projectName} - Project Overview`,
          content: homepageContent.substring(0, 200) + '...'
        };
        
        // Create navigation links
        const navLinks = teamFolders.map(folder => ({
          name: folder,
          url: `${siteUrl}/Shared Documents/${mainFolder.path}/${folder}`
        }));
        
        siteUpdates.navigation = {
          updated: true,
          links: navLinks
        };
        
      } catch (updateError) {
        siteUpdates.error = updateError.message;
      }
    }
    
    res.json({
      success: createdFiles.length > 0,
      filesCreated: createdFiles.length,
      files: createdFiles,
      failedFiles: failedFiles,
      projectFolder: mainFolder.path,
      nestedStructure: createFolders,
      siteUpdates: siteUpdates,
      message: `Created ${createdFiles.length} files in SharePoint${createFolders ? ' with nested folders' : ''}${updateSiteContent ? ' and updated site content' : ''}`
    });
    
  } catch (error) {
    console.error('SharePoint error:', error.response?.data || error.message);
    
    if (error.response?.data?.error?.code === 'InvalidAuthenticationToken') {
      res.status(401).json({ 
        success: false, 
        error: 'Access token expired or invalid. Please re-authenticate with SharePoint.',
        code: 'TOKEN_EXPIRED'
      });
    } else {
      res.status(500).json({ success: false, error: error.response?.data?.error?.message || error.message });
    }
  }
});
// Gmail email generation with attachments
app.post('/api/gmail/generate-emails', async (req, res) => {
  try {
    const { 
      prompt, 
      accessToken, 
      emailCount = 5,
      recipientEmail = 'test@example.com'
    } = req.body;
    
    if (!accessToken || !prompt) {
      return res.status(400).json({ 
        success: false, 
        error: 'Access token and prompt are required' 
      });
    }
    
    const createdEmails = [];
    const failedEmails = [];
    
    for (let i = 0; i < emailCount; i++) {
      try {
        // Generate unique subject
        const subjectPrompt = `Generate a unique professional email subject line for email ${i + 1} related to: "${prompt}". Make it specific and actionable. Return only the subject line:`;
        
        const subjectInput = {
          modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
          contentType: 'application/json',
          accept: 'application/json',
          body: JSON.stringify({
            anthropic_version: 'bedrock-2023-05-31',
            max_tokens: 100,
            messages: [{ role: 'user', content: subjectPrompt }]
          })
        };
        
        const subjectCommand = new InvokeModelCommand(subjectInput);
        const subjectResponse = await client.send(subjectCommand);
        const subjectBody = JSON.parse(new TextDecoder().decode(subjectResponse.body));
        const subject = subjectBody.content[0].text.trim().replace(/"/g, '');
        
        // Generate email body
        const bodyPrompt = `Write a professional email body for: "${subject}" related to project: "${prompt}". Write 150-200 words. Include specific details and actionable items. Return only the email content:`;
        
        const bodyInput = {
          modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
          contentType: 'application/json',
          accept: 'application/json',
          body: JSON.stringify({
            anthropic_version: 'bedrock-2023-05-31',
            max_tokens: 500,
            messages: [{ role: 'user', content: bodyPrompt }]
          })
        };
        
        const bodyCommand = new InvokeModelCommand(bodyInput);
        const bodyResponse = await client.send(bodyCommand);
        const bodyResponseBody = JSON.parse(new TextDecoder().decode(bodyResponse.body));
        const emailBody = bodyResponseBody.content[0].text.trim();
        
        // Create attachment
        const attachmentType = ['TXT', 'CSV', 'HTML'][Math.floor(Math.random() * 3)];
        const attachmentPrompt = `Generate ${attachmentType} content for attachment related to: "${subject}". Write 100-150 words of relevant professional content. Return only the content:`;
        
        const attachmentInput = {
          modelId: 'anthropic.claude-3-5-sonnet-20241022-v2:0',
          contentType: 'application/json',
          accept: 'application/json',
          body: JSON.stringify({
            anthropic_version: 'bedrock-2023-05-31',
            max_tokens: 400,
            messages: [{ role: 'user', content: attachmentPrompt }]
          })
        };
        
        const attachmentCommand = new InvokeModelCommand(attachmentInput);
        const attachmentResponse = await client.send(attachmentCommand);
        const attachmentResponseBody = JSON.parse(new TextDecoder().decode(attachmentResponse.body));
        const attachmentContent = attachmentResponseBody.content[0].text.trim();
        
        // Create attachment file
        const extensions = { TXT: '.txt', CSV: '.csv', HTML: '.html' };
        const attachmentName = `attachment_${i + 1}${extensions[attachmentType]}`;
        const attachmentBase64 = Buffer.from(attachmentContent).toString('base64');
        
        // Compose email
        const email = {
          raw: Buffer.from(
            `To: ${recipientEmail}\r\n` +
            `Subject: ${subject}\r\n` +
            `Content-Type: multipart/mixed; boundary="boundary123"\r\n\r\n` +
            `--boundary123\r\n` +
            `Content-Type: text/plain; charset="UTF-8"\r\n\r\n` +
            `${emailBody}\r\n\r\n` +
            `--boundary123\r\n` +
            `Content-Type: text/plain; name="${attachmentName}"\r\n` +
            `Content-Disposition: attachment; filename="${attachmentName}"\r\n` +
            `Content-Transfer-Encoding: base64\r\n\r\n` +
            `${attachmentBase64}\r\n` +
            `--boundary123--`
          ).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
        };
        
        // Send email via Gmail API
        const emailResponse = await axios.post(
          'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
          email,
          {
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'application/json'
            }
          }
        );
        
        createdEmails.push({
          id: emailResponse.data.id,
          subject: subject,
          recipient: recipientEmail,
          attachmentName: attachmentName,
          attachmentType: attachmentType,
          bodyLength: emailBody.length
        });
        
        // Small delay between emails
        await new Promise(resolve => setTimeout(resolve, 1000));
        
      } catch (emailError) {
        console.error(`Failed to create email ${i + 1}:`, emailError.response?.data || emailError.message);
        failedEmails.push({
          index: i + 1,
          error: emailError.response?.data?.error?.message || emailError.message
        });
      }
    }
    
    res.json({
      success: createdEmails.length > 0,
      emailsCreated: createdEmails.length,
      emails: createdEmails,
      failedEmails: failedEmails,
      message: `Generated ${createdEmails.length} emails with unique subjects and attachments`
    });
    
  } catch (error) {
    console.error('Gmail error:', error.response?.data || error.message);
    
    if (error.response?.data?.error?.code === 'invalid_grant') {
      res.status(401).json({ 
        success: false, 
        error: 'Access token expired or invalid. Please re-authenticate with Gmail.',
        code: 'TOKEN_EXPIRED'
      });
    } else {
      res.status(500).json({ success: false, error: error.response?.data?.error?.message || error.message });
    }
  }
});