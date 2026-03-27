/**
 * Excel EZ - Módulo de Autenticación Local
 * Gestiona el almacenamiento de usuarios en localStorage y la sesión activa en sessionStorage.
 */

const AUTH = {
    // Claves para almacenamiento
    USERS_KEY: 'excelez_usuarios',
    SESSION_KEY: 'sesion_activa',

    // Obtener todos los usuarios
    getUsers: () => {
        const users = localStorage.getItem(AUTH.USERS_KEY);
        return users ? JSON.parse(users) : [];
    },

    // Registrar un nuevo usuario
    register: (nombre, email, password) => {
        const usuarios = AUTH.getUsers();
        
        // Verificar si el email ya existe
        const existe = usuarios.find(u => u.email === email);
        if (existe) {
            return { 
                success: false, 
                errorType: 'email',
                message: 'Este correo ya está registrado. <br><a href="index.html" style="color:var(--primary);text-decoration:none;">¿Quieres iniciar sesión?</a>' 
            };
        }

        const newUser = {
            nombre,
            email,
            password: btoa(password), // codificación básica
            fechaRegistro: new Date().toISOString()
        };

        usuarios.push(newUser);
        localStorage.setItem(AUTH.USERS_KEY, JSON.stringify(usuarios));
        
        return { success: true };
    },

    // Iniciar sesión
    login: (email, password) => {
        const usuarios = AUTH.getUsers();
        const usuario = usuarios.find(u => u.email === email.trim());

        // Verificar si existe
        if (!usuario) {
            return { 
                success: false, 
                errorType: 'email',
                message: 'Este correo no está registrado. <br><a href="register.html" style="color:var(--primary);text-decoration:none;">¿Quieres crear una cuenta?</a>' 
            };
        }

        // Verificar contraseña
        if (usuario.password !== btoa(password)) {
            return { 
                success: false, 
                errorType: 'password',
                message: 'Contraseña incorrecta. Inténtalo de nuevo.' 
            };
        }

        // Login exitoso — guardar sesión activa
        const sessionData = { 
            email: usuario.email, 
            nombre: usuario.nombre,
            fechaLogin: new Date().toISOString()
        };
        sessionStorage.setItem(AUTH.SESSION_KEY, JSON.stringify(sessionData));
        return { success: true };
    },

    // Cerrar sesión
    logout: () => {
        sessionStorage.removeItem(AUTH.SESSION_KEY);
        // NO borrar localStorage (usuarios registrados se mantienen)
        window.location.href = './index.html';
    },

    // Verificar si hay una sesión activa
    checkSession: () => {
        const session = sessionStorage.getItem(AUTH.SESSION_KEY);
        if (!session) {
            window.location.href = './index.html';
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
