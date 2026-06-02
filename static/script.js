// Voice Input Support
const voiceBtns = document.querySelectorAll('.voice-btn');
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

if (SpeechRecognition) {
    const recognition = new SpeechRecognition();
    recognition.continuous = false;
    recognition.interimResults = false;

    voiceBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.getAttribute('data-target');
            const targetInput = document.getElementById(targetId);
            const originalIcon = btn.innerHTML;
            
            btn.innerHTML = '<span class="material-icons align-middle spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>';
            btn.disabled = true;
            
            recognition.lang = document.getElementById('language').value === 'Hindi' ? 'hi-IN' : 'en-US';
            recognition.start();

            recognition.onresult = (event) => {
                const transcript = event.results[0][0].transcript;
                targetInput.value = transcript;
                btn.innerHTML = originalIcon;
                btn.disabled = false;
            };

            recognition.onerror = () => {
                btn.innerHTML = originalIcon;
                btn.disabled = false;
                alert('Voice recognition error. Please try again.');
            };

            recognition.onspeechend = () => {
                recognition.stop();
                btn.innerHTML = originalIcon;
                btn.disabled = false;
            };
        });
    });
} else {
    voiceBtns.forEach(btn => btn.style.display = 'none');
}

document.getElementById('pptForm').addEventListener('submit', async (e) => {
    e.preventDefault();

    const title = document.getElementById('title').value;
    const topic = document.getElementById('topic').value;
    const language = document.getElementById('language').value;
    const template = document.getElementById('template').value;
    const num_slides = 8;
    const format = 'pptx';
    const statusText = document.getElementById('statusText');
    const timer = document.getElementById('timer');
    const downloadButtons = document.getElementById('downloadButtons');
    const downloadLink = document.getElementById('downloadLink');

    statusText.textContent = 'Generating masterpiece...';
    statusText.className = 'text-muted';
    timer.classList.remove('d-none');
    downloadButtons.classList.add('d-none');

    let seconds = 0;
    const timerInterval = setInterval(() => {
        seconds++;
        timer.textContent = `${seconds}s`;
    }, 1000);

    try {
        const response = await fetch('/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ title, topic, language, template, format, num_slides }),
        });

        clearInterval(timerInterval);

        if (!response.ok) {
            const errData = await response.json();
            throw new Error(errData.error || 'Failed to generate presentation');
        }

        const resData = await response.json();

        downloadLink.href = resData.download_url;
        downloadLink.download = resData.download_name;
        downloadButtons.classList.remove('d-none');
        statusText.textContent = 'Success! Your masterpiece is ready.';
        statusText.className = 'text-success';
        timer.classList.add('d-none');

        // Auto-trigger download for mobile compatibility
        window.location.href = resData.download_url;
    } catch (error) {
        clearInterval(timerInterval);
        statusText.textContent = `Error: ${error.message}`;
        statusText.className = 'text-danger';
        timer.classList.add('d-none');
    }
});