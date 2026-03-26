/**
 * Excel EZ - Módulo de Autenticación Local
 * Gestiona el almacenamiento de usuarios en localStorage y la sesión activa en sessionStorage.
 */

const AUTH = {
    // Claves para almacenamiento
    USERS_KEY: 'excelez_users',
    SESSION_KEY: 'excelez_session',

    // Obtener todos los usuarios
    getUsers: () => {
        const users = localStorage.getItem(AUTH.USERS_KEY);
        return users ? JSON.parse(users) : [];
    },

    // Registrar un nuevo usuario
    register: (name, email, password) => {
        const users = AUTH.getUsers();
        
        // Verificar si el email ya existe
        if (users.some(u => u.email === email)) {
            return { success: false, message: 'El correo electrónico ya está registrado.' };
        }

        const newUser = {
            name,
            email,
            password: btoa(password), // Hashing simple con Base64 para propósitos de demostración
            date: new Date().toISOString()
        };

        users.push(newUser);
        localStorage.setItem(AUTH.USERS_KEY, JSON.stringify(users));
        return { success: true };
    },

    // Iniciar sesión
    login: (email, password) => {
        const users = AUTH.getUsers();
        const user = users.find(u => u.email === email && u.password === btoa(password));

        if (user) {
            const sessionData = { 
                email: user.email, 
                name: user.name,
                loginTime: new Date().getTime()
            };
            sessionStorage.setItem(AUTH.SESSION_KEY, JSON.stringify(sessionData));
            return { success: true };
        }

        return { success: false, message: 'Correo o contraseña incorrectos.' };
    },

    // Cerrar sesión
    logout: () => {
        sessionStorage.removeItem(AUTH.SESSION_KEY);
        window.location.href = 'index.html';
    },

    // Verificar si hay una sesión activa
    checkSession: () => {
        const session = sessionStorage.getItem(AUTH.SESSION_KEY);
        if (!session) {
            window.location.href = 'index.html';
            return false;
        }
        return JSON.parse(session);
    },

    // Obtener datos del usuario actual
    getCurrentUser: () => {
        const session = sessionStorage.getItem(AUTH.SESSION_KEY);
        return session ? JSON.parse(session) : null;
    }
};
