document.addEventListener('DOMContentLoaded', () => {
    const loginForm = document.getElementById('loginForm');
    const emailInput = document.getElementById('email');
    const passwordInput = document.getElementById('password');
    const emailError = document.getElementById('emailError');
    const passwordError = document.getElementById('passwordError');
    const togglePasswordBtn = document.getElementById('togglePassword');
    const eyeIcon = document.getElementById('eyeIcon');
    const loginBtn = document.getElementById('loginBtn');
    const btnText = loginBtn.querySelector('.btn-text');
    const btnSpinner = document.getElementById('btnSpinner');
    const successMsg = document.getElementById('successMessage');

    // Check for "registered" parameter
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('registered') === 'true') {
        successMsg.style.display = 'block';
    }

    // Initialize Lucide icons
    lucide.createIcons();

    // Password visibility toggle
    togglePasswordBtn.addEventListener('click', () => {
        const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
        passwordInput.setAttribute('type', type);
        
        // Update icon
        const iconName = type === 'text' ? 'eye-off' : 'eye';
        eyeIcon.setAttribute('data-lucide', iconName);
        lucide.createIcons();
    });

    // Simple email validation
    const isValidEmail = (email) => {
        return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    };

    // Show/Hide error functions
    const showError = (element, message, input) => {
        element.textContent = message;
        element.classList.add('visible');
        input.classList.add('invalid');
    };

    const hideError = (element, input) => {
        element.textContent = '';
        element.classList.remove('visible');
        input.classList.remove('invalid');
    };

    // Real-time validation
    emailInput.addEventListener('input', () => {
        if (emailInput.value.trim() !== '') {
            hideError(emailError, emailInput);
        }
    });

    passwordInput.addEventListener('input', () => {
        if (passwordInput.value.trim() !== '') {
            hideError(passwordError, passwordInput);
        }
    });

    // Form submission
    loginForm.addEventListener('submit', (e) => {
        e.preventDefault();
        
        let isValid = true;
        const email = emailInput.value.trim();
        const password = passwordInput.value.trim();

        // Email validation
        if (!email) {
            showError(emailError, 'El correo electrónico es requerido', emailInput);
            isValid = false;
        } else if (!isValidEmail(email)) {
            showError(emailError, 'Por favor, ingresa un correo válido', emailInput);
            isValid = false;
        }

        // Password validation
        if (!password) {
            showError(passwordError, 'La contraseña es requerida', passwordInput);
            isValid = false;
        }

        if (isValid) {
            // Start loading
            loginBtn.disabled = true;
            btnText.style.opacity = '0';
            btnSpinner.style.display = 'block';
            successMsg.style.display = 'none';
            
            // Artificial delay for UX
            setTimeout(() => {
                const result = AUTH.login(email, password);
                
                if (result.success) {
                    window.location.href = 'dashboard.html';
                } else {
                    showError(passwordError, result.message, passwordInput);
                    loginBtn.disabled = false;
                    btnText.style.opacity = '1';
                    btnSpinner.style.display = 'none';
                }
            }, 1200);
        }
    });
});
