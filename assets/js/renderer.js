const { dialog } = require('electron');

function openFileDialog(fileNumber) {
    // Check for the various File API support.
    if (window.File && window.FileReader && window.FileList && window.Blob) {
  
      // Create a file input element
      var fileInput = document.createElement("input");
  
      // Set its type to file
      fileInput.type = "file";
  
      // Add an onchange event listener to the file input element
      fileInput.addEventListener("change", function() {
        // Get the selected file
        var file = this.files[0];
  
        // Get the file path
        var filePath = file.path;
        var fileName = file.name;
  
        // Update the p element with the selected file path
        document.getElementById(fileNumber + "-selected").innerHTML = "File Name: " + fileName;
      });
  
      // Simulate a click on the file input element
      fileInput.click();
    } else {
      alert("The File APIs are not fully supported in this browser.");
    }
  }