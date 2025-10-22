"use strict";

require("dotenv").config();
const express = require("express");
const session = require("express-session");
const multer = require("multer");
const axios = require("axios");
const fs = require("fs");
const path = require("path");
const querystring = require("querystring");

const app = express();
const upload = multer({ dest: "uploads/" });

app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));
app.use(session({ secret: "secret", resave: false, saveUninitialized: true }));

const PORT = 5000; // changed
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI;
const SCOPES = process.env.SCOPES;

// --- Login route ---
app.get("/auth/login", (req, res) => {
  const authUrl = `https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?${querystring.stringify(
    {
      client_id: CLIENT_ID,
      response_type: "code",
      redirect_uri: REDIRECT_URI,
      response_mode: "query",
      scope: SCOPES,
      state: "12345",
    }
  )}`;
  res.redirect(authUrl);
});

// --- OAuth callback ---
app.get("/auth/callback", async (req, res) => {
  const code = req.query.code;
  if (!code) return res.send("No code received");

  try {
    const tokenRes = await axios.post(
      "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
      querystring.stringify({
        client_id: CLIENT_ID,
        scope: SCOPES,
        code,
        redirect_uri: REDIRECT_URI,
        grant_type: "authorization_code",
        client_secret: CLIENT_SECRET,
      }),
      { headers: { "Content-Type": "application/x-www-form-urlencoded" } }
    );

    const accessToken = tokenRes.data.access_token;
    // Lấy thông tin user
    const userRes = await axios.get("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    req.session.accessToken = accessToken;
    req.session.user = userRes.data;
    res.redirect("/");
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.send("Error during authentication");
  }
});

// --- Logout ---
app.get("/logout", (req, res) => {
  req.session.destroy();
  res.redirect("/");
});

// --- Session info (for HTML JS) ---
app.get("/session-info", (req, res) => {
  res.json({ loggedIn: !!req.session.user, user: req.session.user || null });
});

// --- Serve HTML ---
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public/index.html"));
});

// --- Convert DOCX to PDF ---
app.post("/convert", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).send("File required.");
  if (!req.session.accessToken)
    return res.status(401).send("Not authenticated.");

  const filePath = path.resolve(req.file.path);
  const fileName = req.file.originalname;
  const fileData = fs.readFileSync(filePath);

  try {
    // Upload to user's OneDrive root
    const uploadRes = await axios.put(
      `https://graph.microsoft.com/v1.0/me/drive/root:/${fileName}:/content`,
      fileData,
      {
        headers: {
          Authorization: `Bearer ${req.session.accessToken}`,
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        },
      }
    );

    const itemId = uploadRes.data.id;

    // Convert to PDF
    const pdfRes = await axios.get(
      `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content?format=pdf`,
      {
        headers: { Authorization: `Bearer ${req.session.accessToken}` },
        responseType: "arraybuffer",
      }
    );

    res.set({
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="${fileName.replace(
        /\.docx$/,
        ".pdf"
      )}"`,
    });
    res.send(pdfRes.data);

    fs.unlinkSync(filePath);
    await axios.delete(
      `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}`,
      {
        headers: { Authorization: `Bearer ${req.session.accessToken}` },
      }
    );
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).send("Error converting file.");
  }
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
