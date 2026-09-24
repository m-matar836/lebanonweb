// ===================================================================
//                      3. منطق صفحة تسجيل الدخول
// ===================================================================
async function handleLoginPage() {
    document.getElementById('loginForm').addEventListener('submit', async (e) => {
        e.preventDefault();
        const username = e.target.username.value.trim().toLowerCase();
        const password = e.target.password.value.trim();
        const rememberMe = e.target.rememberMe.checked;
        const submitBtn = e.target.querySelector('button[type="submit"]');
        const errorMessage = document.getElementById('errorMessage');
        // تسريع الدخول: نجلب البيانات الأساسية بالتوازي مع تحقق الهوية (بدل التسلسل).
        const dbFetchPromise = (typeof getDbData === 'function') ? getDbData().catch(() => null) : Promise.resolve(null);
        submitBtn.disabled = true;
        submitBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> جار التحقق...';
        errorMessage.textContent = '';
        try {
            const loginResult = await apiPost('doLogin', { username, password });
            if (loginResult.status !== 'success') throw new Error('Invalid credentials');
            
            if (rememberMe) {
                localStorage.setItem('currentUser', JSON.stringify(loginResult.user));
                localStorage.setItem('loginTimestamp', Date.now());
            } else {
                sessionStorage.setItem('currentUser', JSON.stringify(loginResult.user));
                sessionStorage.setItem('loginTimestamp', Date.now());
            }
            
            await dbFetchPromise;
            errorMessage.textContent = 'تم التحقق بنجاح! جارٍ التحويل...';
            errorMessage.style.color = '#2ecc71';
            // V50: تسجيل الجلسة في دفتر الجلسات المحلي — مع إشعار عند الدخول من جهاز جديد.
            const isNewDevice = recordLoginSession(loginResult.user?.name || username);
            if (isNewDevice) {
                const notice = document.createElement('div');
                notice.className = 'alert alert-warning text-center small mt-2';
                notice.innerHTML = '<i class="fa-solid fa-shield-halved me-1"></i> تم تسجيل الدخول من جهاز جديد على هذا الحساب. إن لم تكن أنت، فغيّر كلمة المرور فوراً.';
                const card = document.querySelector('#view-login .login-form');
                if (card) {
                    const old = card.querySelector('.session-device-notice');
                    if (old) old.remove();
                    notice.classList.add('session-device-notice');
                    card.appendChild(notice);
                }
            }
            // V45/V46: حساب "user" ومنصبه "مروج" ينتقل للدوام، بقية الأدوار لشاشة التقارير.
            setTimeout(() => { navigateTo(isPromoterAccount(loginResult.user) ? 'attendance' : 'reports'); }, 1000);
        } catch (error) {
            errorMessage.textContent = 'اسم المستخدم أو كلمة المرور غير صحيحة.';
            submitBtn.disabled = false;
            submitBtn.innerHTML = 'دخـــول';
        }
    });

    const togglePassword = document.querySelector('.toggle-password');
    if(togglePassword) {
        togglePassword.addEventListener('click', function () {
            const passwordInput = document.getElementById('password');
            const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
            passwordInput.setAttribute('type', type);
            this.classList.toggle('fa-eye');
            this.classList.toggle('fa-eye-slash');
        });
    }
}

// V58: تسجيل شاشة الدخول في موجه الـ SPA — تُنفَّذ handleLoginPage عند أول فتح.
if (typeof registerView === 'function') registerView('login', handleLoginPage);

