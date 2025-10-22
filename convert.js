const axios = require("axios");

/**
 * Convert DOCX → PDF dùng Microsoft Graph
 * @param {Buffer} docxBuffer
 * @param {string} fileName
 * @param {string} accessToken
 * @returns {Buffer} PDF
 */
async function convertDocxToPdf(docxBuffer, fileName, accessToken) {
  const safeName = fileName.replace(/\s+/g, "_");
  const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(
    safeName
  )}:/content`;

  // 🔹 Upload DOCX tạm
  const uploadRes = await axios.put(uploadUrl, docxBuffer, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    },
  });

  const fileId = uploadRes.data.id;

  // 🔹 Convert sang PDF
  const convertUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/content?format=pdf`;
  const pdfRes = await axios.get(convertUrl, {
    headers: { Authorization: `Bearer ${accessToken}` },
    responseType: "arraybuffer",
  });

  // 🔹 Xoá DOCX tạm
  try {
    await axios.delete(
      `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );
  } catch {}

  return Buffer.from(pdfRes.data);
}

module.exports = { convertDocxToPdf };
