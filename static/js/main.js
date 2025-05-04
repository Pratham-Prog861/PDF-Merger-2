document.addEventListener("DOMContentLoaded", () => {
    const preloader = document.getElementById("preloader");
    const progress = document.getElementById("progress");
    const mainContent = document.getElementById("mainContent");
    let progressValue = 0;

    const interval = setInterval(() => {
      progressValue += 1;
      progress.textContent = `${progressValue}%`;

      if (progressValue >= 100) {
        clearInterval(interval);
        setTimeout(() => {
          preloader.classList.add("fade-out");
          mainContent.classList.add("visible");
        }, 500);
      }
    }, 20);
  });

// Initialize Lenis with smooth scrolling options
const lenis = new Lenis({
  duration: 1.2,
  easing: (t) => Math.min(1, 1.001 - Math.pow(2, -10 * t)),
  direction: 'vertical',
  gestureDirection: 'vertical',
  smooth: true,
  smoothTouch: false,
  touchMultiplier: 2
});

function raf(time) {
  lenis.raf(time);
  requestAnimationFrame(raf);
}

requestAnimationFrame(raf);

  // Add smooth scroll to all anchor links
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
      e.preventDefault();
      const targetId = this.getAttribute('href');
      const targetElement = document.querySelector(targetId);
      
      if (targetElement) {
        lenis.scrollTo(targetElement, {
          offset: -100, // Adjust offset to account for fixed header
          duration: 1.2
        });
      }
    });
  });

// Handle form submissions
document.getElementById('mergeForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const formData = new FormData();
    const files = document.getElementById('mergeFiles').files;
    
    for (let file of files) {
        formData.append('files[]', file);
    }
    
    try {
        const response = await fetch('/merge', {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'merged.pdf';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            showSuccess('PDFs merged successfully!');
        } else {
            const error = await response.json();
            showError(error.error || 'Failed to merge PDFs');
        }
    } catch (error) {
        showError('An error occurred while merging PDFs');
    }
});

// Show success message
function showSuccess(message) {
    const alert = document.getElementById('successAlert');
    const messageSpan = document.getElementById('successMessage');
    messageSpan.textContent = message;
    alert.classList.remove('d-none');
    setTimeout(() => {
        alert.classList.add('d-none');
    }, 5000);
}

// Show error message
function showError(message) {
    const alert = document.getElementById('successAlert');
    const messageSpan = document.getElementById('successMessage');
    alert.classList.remove('alert-success');
    alert.classList.add('alert-danger');
    messageSpan.textContent = message;
    alert.classList.remove('d-none');
    setTimeout(() => {
        alert.classList.add('d-none');
        alert.classList.remove('alert-danger');
        alert.classList.add('alert-success');
    }, 5000);
};

// Add this after your existing code
// Add compression form handler
document.getElementById('compressForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const formData = new FormData();
    const file = document.getElementById('compressFile').files[0];
    const level = document.getElementById('compressionLevel').value;
    
    if (!file) {
        showError('Please select a PDF file');
        return;
    }
    
    formData.append('file', file);
    formData.append('level', level);
    
    try {
        const response = await fetch('/compress', {
            method: 'POST',
            body: formData
        });
        
        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = 'compressed.pdf';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            a.remove();
            showSuccess('PDF compressed successfully!');
        } else {
            const error = await response.json();
            showError(error.error || 'Failed to compress PDF');
        }
    } catch (error) {
        showError('An error occurred while compressing PDF');
        console.error('Compression error:', error);
    }
});

// Update compression level display
document.getElementById('compressionLevel')?.addEventListener('input', (e) => {
    const valueDisplay = document.getElementById('compressionValue');
    if (valueDisplay) {
        valueDisplay.textContent = `${e.target.value}%`;
    }
});

// Update PPT to PDF conversion handler
document.getElementById('convertForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const formData = new FormData();
    const file = document.getElementById('pptFile').files[0];
    
    if (!file) {
        showError('Please select a PowerPoint file');
        return;
    }
    
    // Validate file type
    if (!file.name.match(/\.(ppt|pptx)$/i)) {
        showError('Please select a valid PowerPoint file (.ppt or .pptx)');
        return;
    }
    
    // Show loading state
    const submitButton = e.target.querySelector('button[type="submit"]');
    const originalText = submitButton.innerHTML;
    submitButton.innerHTML = '<i class="fas fa-spinner fa-spin mr-2"></i>Converting...';
    submitButton.disabled = true;
    
    formData.append('file', file);
    
    try {
        const response = await fetch('/convert-ppt', {
            method: 'POST',
            body: formData,
            headers: {
                'Accept': 'application/pdf'
            }
        });
        
        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Conversion failed');
        }
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = file.name.replace(/\.(ppt|pptx)$/i, '.pdf');
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        showSuccess('PowerPoint file converted successfully!');
    } catch (error) {
        console.error('Conversion error:', error);
        showError(error.message || 'Failed to convert PowerPoint file');
    } finally {
        submitButton.innerHTML = originalText;
        submitButton.disabled = false;
    }
});