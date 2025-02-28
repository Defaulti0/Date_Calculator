Office.onReady(() => {
  document.getElementById("updateButton").onclick = updateApptDate;
  applyOfficeTheme();
});

const currentDate = new Date();
const monthYear = document.getElementById("monthYear");
const calendarGrid = document.getElementById("calendarGrid");

function generateCalendar(date) {
  const year = date.getFullYear();
  const month = date.getMonth();
  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  monthYear.textContent = date.toLocaleString("default", { month: "long", year: "numeric" });
  calendarGrid.innerHTML = "";

  const weekdays = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  weekdays.forEach((day) => {
    const dayElement = document.createElement("div");
    dayElement.classList.add("calendar-day", "weekday-header");
    dayElement.textContent = day;
    calendarGrid.appendChild(dayElement);
  });

  for (let i = 0; i < firstDay; i++) {
    const emptyDay = document.createElement("div");
    calendarGrid.appendChild(emptyDay);
  }

  for (let day = 1; day <= daysInMonth; day++) {
    const dayElement = document.createElement("div");
    dayElement.classList.add("calendar-day");
    dayElement.textContent = day;

    // Use the passed 'date' object for highlighting
    if (year === date.getFullYear() && month === date.getMonth() && day === date.getDate()) {
      dayElement.classList.add("today");
    }

    calendarGrid.appendChild(dayElement);
  }
}

function applyOfficeTheme() {
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  console.log(bodyBackgroundColor);

  const calendarDays = document.querySelectorAll(".calendar-day");
  const weekdayHeaders = document.querySelectorAll(".calendar-day.weekday-header");

  if (bodyBackgroundColor == "#FAF9F8") {
    // Apply body background color to a CSS class.
    document.getElementById("sampleList").style.color = "black";
    calendarDays.forEach((day) => {
      day.style.backgroundColor = "#white";
      day.style.color = "black";
    });
    weekdayHeaders.forEach((day) => {
      day.style.backgroundColor = "#white";
      day.style.color = "black";
    });
  } else if (bodyBackgroundColor == "#212121") {
    // Apply body background color to a CSS class.
    document.getElementById("wholePage").style.color = "white";
    calendarDays.forEach((day) => {
      day.style.backgroundColor = "#black";
      day.style.color = "white";
    });
    weekdayHeaders.forEach((day) => {
      day.style.backgroundColor = "#black";
      day.style.color = "white";
    });
  }
}

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
        generateCalendar(parsedDate);
        // updateApptTime();
      }
    });
  } catch (error) {
    alert("Error updating appointment: " + error);
  }
}

// function updateApptTime(ogDate) {
//   const startTimeValue = document.getElementById("startTimePicker").value;
//   const endTimeValue = document.getElementById("endTimePicker").value;

//   const originalStart = Office.context.mailbox.item.start;
//   const originalEnd = Office.context.mailbox.item.end;

//   const startDateTime = new Date(originalStart); // Create a copy of original start
//   const endDateTime = new Date(originalEnd); // Create a copy of original end

//   startDateTime.setHours(parseInt(startTimeValue.split(":")[0]));
//   startDateTime.setMinutes(parseInt(startTimeValue.split(":")[1]));

//   endDateTime.setHours(parseInt(endTimeValue.split(":")[0]));
//   endDateTime.setMinutes(parseInt(endTimeValue.split(":")[1]));

//   Office.context.mailbox.item.start.setAsync(startDateTime, (startResult) => {
//     if (startResult.status === Office.AsyncResultStatus.Succeeded) {
//       Office.context.mailbox.item.end.setAsync(endDateTime, (endResult) => {
//         if (endResult.status !== Office.AsyncResultStatus.Succeeded) {
//           console.error("error setting end time: ", endResult.error);
//         }
//       });
//     } else {
//       console.error("error setting start time: ", startResult.error);
//     }
//   });
// }

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

// function populateTimePicker(elementId) {
//   const timePicker = document.getElementById(elementId);
//   timePicker.innerHTML = "";

//   const startTime = new Date();
//   startTime.setHours(0, 0, 0, 0);

//   const endTime = new Date();
//   endTime.setHours(23, 59, 0, 0);

//   const interval = 30 * 60 * 1000;

//   let currentTime = new Date(startTime);

//   while (currentTime <= endTime) {
//     const hours = currentTime.getHours().toString().padStart(2, "0");
//     const minutes = currentTime.getMinutes().toString().padStart(2, "0");
//     const timeString = `${hours}:${minutes}`;

//     const option = document.createElement("option");
//     option.value = timeString;
//     option.textContent = timeString;
//     timePicker.appendChild(option);

//     currentTime.setTime(currentTime.getTime() + interval);
//   }
// }

document.addEventListener("DOMContentLoaded", function () {
  const button = document.getElementById("toggleButton"); // Replace with the actual ID
  if (button) {
    button.addEventListener("click", function () {
      console.log("Button clicked!");
      if (myList.style.display === "none" || myList.style.display === "") {
        myList.style.display = "block"; // Or 'list-item'
        toggleButton.textContent = "Hide Examples";
      } else {
        myList.style.display = "none";
        toggleButton.textContent = "Show Examples";
      }
    });
  } else {
    console.error("Button not found!");
  }
});

// populateTimePicker("startTimePicker");
// populateTimePicker("endTimePicker");
generateCalendar(currentDate);
