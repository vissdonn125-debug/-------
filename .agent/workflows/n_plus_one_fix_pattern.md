---
description: Pattern to fix N+1 performance issues in Google Apps Script
---

# Fix N+1 Query Pattern

This skill documents the pattern to resolve N+1 performance issues where `google.script.run` is called in a loop.

## Problem
Calling `google.script.run` inside a loop (e.g., `forEach`) causes significant performance degradation because each call is a separate network request.

**Bad Code:**
```javascript
list.forEach(item => {
  // ⛔️ Bad: Triggers server call for every item
  google.script.run.withSuccessHandler(img => {
     render(img);
  }).getImage(item.id);
});
```

## Solution Pattern: Batch Retrieval

### 1. Backend (Server-side)
Create a function that accepts an array of IDs and returns a Map of results.

```javascript
// admin_server.js / server_common.js
function api_getImagesBatch(fileIds) {
  const result = {};
  fileIds.forEach(id => {
    try {
      const file = DriveApp.getFileById(id);
      const blob = file.getBlob();
      result[id] = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      result[id] = null; // Handle error gracefully
    }
  });
  return result;
}
```

### 2. Frontend (Client-side)
1. Collect all IDs first.
2. Render placeholders.
3. Call the batch API once.
4. Update placeholders with results.

```javascript
// 1. Collect IDs
const allIds = list.map(item => item.fileId).filter(id => id);

// 2. Render UI with placeholders (using data-attributes)
list.forEach(item => {
   const html = `<img data-file-id="${item.fileId}" src="placeholder.png">`;
   document.getElementById('container').insertAdjacentHTML('beforeend', html);
});

// 3. Batch Call
if (allIds.length > 0) {
  google.script.run.withSuccessHandler(imgMap => {
    // 4. Update DOM
    Object.keys(imgMap).forEach(id => {
       const b64 = imgMap[id];
       const targets = document.querySelectorAll(`img[data-file-id="${id}"]`);
       targets.forEach(img => img.src = b64);
    });
  }).api_getImagesBatch(allIds);
}
```

## Checklist
- [ ] Backend: Created batch function accepting array.
- [ ] Frontend: Collected IDs into an array.
- [ ] Frontend: Added `data-id` attributes to DOM elements.
- [ ] Frontend: Implemented single `google.script.run` call.
- [ ] Frontend: Callback updates all matching DOM elements.
