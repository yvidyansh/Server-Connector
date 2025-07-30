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
  origin:'*',
  credentials: true
}));
app.use(express.json());

// TODO: FRIEND'S KATAL APP - Replace with actual Katal build path when available
// const katalBuildPath = path.join(__dirname, '../../ConnectorSampleDataGenerator/src/ConnectorSampleDataGeneratorWebsite/build');
// app.use('/katal', express.static(katalBuildPath));

// Serve original React build files
const clientBuildPath = path.join(__dirname, '../client/build');
if (fs.existsSync(clientBuildPath)) {
  app.use('/client', express.static(clientBuildPath));
}

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

const client = new BedrockRuntimeClient({
  region: process.env.AWS_REGION || 'us-west-2',
  credentials: {
    accessKeyId: process.env.AWS_ACCESS_KEY_ID,
    secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY
  }
});

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
    const { prompt, jiraConfig, issueCount = 3, projectName, projectKey } = req.body;
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
            issuetype: { id: defaultIssueType.id }
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
    if (error.message.includes('Invalid character in header')) {
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
  app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
    console.log(`Original client: http://localhost:${PORT}/client`);
    console.log(`APIs available at: http://localhost:${PORT}/api/*`);
  });
}
