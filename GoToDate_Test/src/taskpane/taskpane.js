Office.onReady(() => {
  document.getElementById("updateButton").onclick = updateApptDate;
});

async function updateApptDate() {
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

  // updateApptTime();
}

function updateApptTime() {
  const startTimeValue = document.getElementById('startTimePicker').value;
  const endTimeValue = document.getElementById('endTimePicker').value;

  if (!startTimeValue || !endTimeValue) {
    console.error("Please select start and end times.");
    return;
  }

  Office.context.mailbox.item.getSelectedPropertiesAsync(["start", "end"], (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const appointmentStart = result.value.start;
      const appointmentEnd = result.value.end;

      if (!appointmentStart || !appointmentEnd) {
        console.error("Could not retrieve appointment start/end times.");
        return;
      }

      try {
        const [startHours, startMinutes] = startTimeValue.split(':').map(Number);
        const [endHours, endMinutes] = endTimeValue.split(':').map(Number);

        const newStart = new Date(appointmentStart);
        newStart.setHours(startHours, startMinutes, 0, 0);

        const newEnd = new Date(appointmentEnd);
        newEnd.setHours(endHours, endMinutes, 0, 0);

        Office.context.mailbox.item.start.setAsync(newStart, (startResult) => {
          if (startResult.status === Office.AsyncResultStatus.Succeeded) {
            Office.context.mailbox.item.end.setAsync(newEnd, (endResult) => {
              if (endResult.status === Office.AsyncResultStatus.Succeeded) {
                console.log("Appointment times updated successfully.");
              } else {
                console.error("Failed to update end time:", endResult.error);
              }
            });
          } else {
            console.error("Failed to update start time:", startResult.error);
          }
        });

      } catch (error) {
        console.error("Error updating appointment times:", error);
      }

    } else {
      console.error("Failed to get selected appointment properties:", result.error);
    }
  });
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
