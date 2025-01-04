document.addEventListener("DOMContentLoaded", function () {
    // Typing Effect
    const words = ["Web Developer", "Web Designer", "Frontend Developer"];
    let wordIndex = 0;
    let charIndex = 0;
    let currentWord = '';
    const typingSpeed = 100;
    const erasingSpeed = 50;
    const newWordDelay = 2000;

    // Type function to add characters
    function type() {
        if (charIndex < words[wordIndex].length) {
            currentWord += words[wordIndex].charAt(charIndex);
            document.querySelector('.typing-animation').textContent = currentWord;
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
            document.querySelector('.typing-animation').textContent = currentWord;
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
    const progressBars = document.querySelectorAll('.progress-done');
    
    progressBars.forEach(bar => {
        setTimeout(() => {
            bar.style.width = bar.getAttribute('data-done') + '%';
            bar.style.opacity = 1;
        }, 500);
    });

    // Animate circular skills
    const circles = document.querySelectorAll('.circle');
    
    circles.forEach(circle => {
        const percent = circle.getAttribute('data-percent');
        // Set CSS variable for circular skill percentage
        circle.style.setProperty('--percent', `${percent}%`);
    });
});
