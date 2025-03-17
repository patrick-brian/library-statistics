// Function to close the side tab
function closeSideTab() {
  document.querySelector('.side-tab').style.display = 'none';
}

// Function to change the content based on the clicked module
function changeContent(module) {
  const contentArea = document.getElementById("content-area");
  
  // Define the content for each module
  let content = "";
  switch (module) {
    case 'module1':
      content = "<h2>Module 1 Content</h2><p>This is the content for Module 1.</p>";
      break;
    case 'module2':
      content = "<h2>Module 2 Content</h2><p>This is the content for Module 2.</p>";
      break;
    case 'module3':
      content = "<h2>Module 3 Content</h2><p>This is the content for Module 3.</p>";
      break;
    case 'module4':
      content = "<h2>Module 4 Content</h2><p>This is the content for Module 4.</p>";
      break;
    case 'module5':
      content = "<h2>Module 5 Content</h2><p>This is the content for Module 5.</p>";
      break;
    default:
      content = "<h2>Welcome! Please select a module.</h2>";
      break;
  }
  
  // Update the content area with the selected module's content
  contentArea.innerHTML = content;
}