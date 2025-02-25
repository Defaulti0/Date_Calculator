Office.onReady(() => {
  document.getElementById("updateButton").onclick = updateAppointmentDate;
});

async function updateAppointmentDate() {
  let userInput = document.getElementById("dateInput").value;

  if (!userInput) {
    alert("Please enter a date.");
    return;
  }

  try {
    // Use a date parser (e.g., chrono-node) to process natural language
    let parsedDate = parseDate(userInput);

    if (!parsedDate) {
      alert("Invalid date. Try something like 'next Monday' or 'March 3'.");
      return;
    }

    // parsedDate.toISOString

    // Get the current appointment item
    Office.context.mailbox.item.start.setAsync(parsedDate, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error(asyncResult.error.message);
      } else {
        console.log("Date set to " + parsedDate);
      }
    });
  } catch (error) {
    alert("Error updating appointment: " + error);
  }
}

// Function to parse natural language dates
function parseDate(input) {
  try {
    let chrono = require("chrono-node");
    return chrono.parseDate(input);
  } catch (error) {
    console.error("Error parsing date:", error);
    return null;
  }
}
