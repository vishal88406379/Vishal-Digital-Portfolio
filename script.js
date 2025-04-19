document.addEventListener("DOMContentLoaded", function () {
  // Typing Effect
  const words = ["Web Developer", "Web Designer", "Frontend Developer"];
  let wordIndex = 0;
  let charIndex = 0;
  let currentWord = "";
  const typingSpeed = 100;
  const erasingSpeed = 50;
  const newWordDelay = 2000;

  // Type function to add characters
  function type() {
    if (charIndex < words[wordIndex].length) {
      currentWord += words[wordIndex].charAt(charIndex);
      document.querySelector(".typing-animation").textContent = currentWord;
      charIndex++;
      setTimeout(type, typingSpeed);
    } else {
      // Start erasing after typing completes
      setTimeout(erase, newWordDelay);
    }
  }

  // Erase function to remove characters
  function erase() {
    if (charIndex > 0) {
      currentWord = currentWord.slice(0, -1);
      document.querySelector(".typing-animation").textContent = currentWord;
      charIndex--;
      setTimeout(erase, erasingSpeed);
    } else {
      // Move to next word
      wordIndex = (wordIndex + 1) % words.length;
      setTimeout(type, typingSpeed + 1100);
    }
  }

  // Start typing animation
  type();

  // Animate progress bars
  const progressBars = document.querySelectorAll(".progress-done");

  progressBars.forEach((bar) => {
    setTimeout(() => {
      bar.style.width = bar.getAttribute("data-done") + "%";
      bar.style.opacity = 1;
    }, 500);
  });

  // Animate circular skills
  const circles = document.querySelectorAll(".circle");

  circles.forEach((circle) => {
    const percent = circle.getAttribute("data-percent");
    // Set CSS variable for circular skill percentage
    circle.style.setProperty("--percent", `${percent}%`);
  });

  // Resume upload functionality
  const resumeUploadForm = document.getElementById("resume-upload-form");
  const resumeFile = document.getElementById("resume-file");
  const uploadResumeBtn = document.getElementById("upload-resume-btn");
  const uploadMessage = document.getElementById("upload-message");
  const contactForm = document.getElementById("contact-form");
  const toggleResumeUploadBtn = document.getElementById(
    "toggle-resume-upload"
  );
  const resumeUploadArea = document.getElementById("resume-upload-area");

  uploadResumeBtn.addEventListener("click", function () {
    resumeFile.click();
  });

  resumeFile.addEventListener("change", function () {
    const file = resumeFile.files[0];
    if (file && file.type === "application/pdf") {
      const formData = new FormData(resumeUploadForm);
      fetch("/upload_resume", {
        method: "POST",
        body: formData,
      })
        .then((response) => {
          if (response.ok) {
            uploadMessage.style.display = "block";
            setTimeout(() => {
              uploadMessage.style.display = "none";
            }, 3000);
          } else {
            alert("Failed to upload resume.");
          }
        })
        .catch((error) => {
          console.error("Error:", error);
          alert("An error occurred while uploading the resume.");
        });
    } else {
      alert("Please upload a PDF file.");
    }
  });

  contactForm.addEventListener("submit", function (event) {
    event.preventDefault();
    const name = contactForm.name.value;
    const email = contactForm.email.value;
    const message = contactForm.message.value;
    const subject = contactForm.subject.value;

    fetch("/send_email", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ name, email, message, subject }),
    })
      .then((response) => {
        if (response.ok) {
          alert("Email sent successfully!");
          contactForm.reset();
        } else {
          alert("Failed to send email.");
        }
      })
      .catch((error) => {
        console.error("Error:", error);
        alert("An error occurred while sending the email.");
      });
  });
  function createTreeNode(item, parent) {
    
  }

  function loadDataTree() {
    
  }
  
  loadDataTree();
  toggleResumeUploadBtn.addEventListener("click", function () {
    if (resumeUploadArea.style.display === "none") {
      resumeUploadArea.style.display = "block";
    } else {
      resumeUploadArea.style.display = "none";
    }
  });
});

