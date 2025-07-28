# Katal Integration TODO

## Steps to integrate friend's Katal app:

### 1. Get Katal Build
- Get the `build` folder from friend's laptop after running `npm run build`
- Place it at: `ConnectorSampleDataGenerator/src/ConnectorSampleDataGeneratorWebsite/build/`

### 2. Update server.js
Uncomment these lines:
```javascript
// Line ~25: Uncomment Katal static serving
const katalBuildPath = path.join(__dirname, '../../ConnectorSampleDataGenerator/src/ConnectorSampleDataGeneratorWebsite/build');
app.use('/katal', express.static(katalBuildPath));

// Line ~40: Uncomment Katal API endpoints
app.get('/api/katal-data', (req, res) => {
  res.json({ message: 'Katal data endpoint' });
});

// Line ~280: Uncomment Katal routes
app.get('/katal/*', (req, res) => {
  res.sendFile(path.join(katalBuildPath, 'index.html'));
});
```

### 3. Update API calls in Katal app
Friend needs to update API calls to point to:
- Local: `http://localhost:5000/api`
- Production: Your Lambda endpoint

### 4. CORS Origins
Add friend's dev server to CORS if needed:
```javascript
origin: ['http://localhost:3000', 'http://localhost:4321', 'http://friend-ip:4321']
```

### 5. Test Integration
1. Start this server: `npm run dev`
2. Friend starts Katal dev: `npm run server`
3. Test API calls between apps