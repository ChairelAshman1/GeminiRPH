const express = require('express');
const path = require('path');
const app = express();

// Serve static files from the Gemini Extension folder (existing project content)
const staticFolder = path.join(__dirname, 'Gemini Extension Assistant');
app.use(express.static(staticFolder));

// Fallback to geminiRPH.html for any unmatched route
app.get('*', (req, res) => {
  res.sendFile(path.join(staticFolder, 'geminiRPH.html'));
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});