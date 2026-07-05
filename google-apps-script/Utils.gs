function parsePayload(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  const contents = e.postData.contents;

  // Attempt to parse JSON payload regardless of Content-Type (helps bypass CORS preflight)
  try {
    return JSON.parse(contents);
  } catch (err) {
    // Fall back to URL-encoded parameters if parsing fails
    return e.parameter || {};
  }
}
