Office.onReady(() => {
  document.getElementById("updateButton").onclick = updateApptDate;
  applyOfficeTheme();
});

const monthYear = document.getElementById("monthYear");
const calendarGrid = document.getElementById("calendarGrid");

function generateCalendar(date) {
  const year = date.getFullYear();
  const month = date.getMonth();
  const firstDay = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const today = new Date();

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

const currentDate = new Date();
generateCalendar(currentDate);

function applyOfficeTheme() {
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  console.log(bodyBackgroundColor);

  const calendarDays = document.querySelectorAll(".calendar-day");
  const weekdayHeaders = document.querySelectorAll(".calendar-day.weekday-header");

  if (bodyBackgroundColor == "#FAF9F8") {
    // Apply body background color to a CSS class.
    document.getElementById("sampleList").style.color = "black";
    document.getElementById("title").style.color = "black";
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
    document.getElementById("title").style.color = "white";
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
        updateApptTime();
      }
    });
  } catch (error) {
    alert("Error updating appointment: " + error);
  }
}

function updateApptTime() {
  const startTimeValue = document.getElementById("startTimePicker").value;
  const endTimeValue = document.getElementById("endTimePicker").value;

  // ... (time format validation) ...

  Office.context.mailbox.item.getAsync("start", (startResult) => {
    if (startResult.status === Office.AsyncResultStatus.Succeeded) {
      const appointmentStart = startResult.value;
      Office.context.mailbox.item.getAsync("end", (endResult) => {
        if (endResult.status === Office.AsyncResultStatus.Succeeded) {
          const appointmentEnd = endResult.value;

          try {
            const [startHours, startMinutes] = startTimeValue.split(":").map(Number);
            const [endHours, endMinutes] = endTimeValue.split(":").map(Number);

            const newStart = new Date(appointmentStart);
            newStart.setHours(startHours, startMinutes, 0, 0);

            const newEnd = new Date(appointmentEnd);
            newEnd.setHours(endHours, endMinutes, 0, 0);

            // ... (time range check and setAsync calls) ...
          } catch (error) {
            // ... (error handling) ...
          }
        } else {
          console.error("Failed to get appointment end:", endResult.error);
          alert("Failed to get appointment end: " + endResult.error.message);
        }
      });
    } else {
      console.error("Failed to get appointment start:", startResult.error);
      alert("Failed to get appointment start: " + startResult.error.message);
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
