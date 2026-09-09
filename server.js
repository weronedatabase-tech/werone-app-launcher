const express = require('express');
const path = require('path');

const app = express();
const port = 3000;

// Serve all static files in the root directory
app.use(express.static(path.join(__dirname, '/')));

// For SPA routing fallback (in case it uses History API)
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(port, '0.0.0.0', () => {
  console.log(`Server listening on port ${port} (0.0.0.0)`);
});
