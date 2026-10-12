JEFE_CREDENCIALES = {
    "jcamacho@dinadeco.go.cr": {
        "nombre": "Jenier Camacho Picado",
        "rol": "ADMIN",
        "pin": "2026"
    },
    "ggutierrez@dinadeco.go.cr": {
        "nombre": "Gabriela Gutiérrez Araya",
        "rol": "ADMIN",
        "pin": "5678"
    }
}

SMTP_CONFIG = {
    "server": "smtp.office365.com",
    "port": 587,
    "user": "bodega@dinadeco.go.cr",  # Correo desde donde sale la notificación
    "password": "TU_CONTRASEÑA_AQUI", # Contraseña o contraseña de aplicación
    "destinatario_principal": "agonzalez@dinadeco.go.cr",
    "destinatarios_cc": [
        "ggutierrez@dinadeco.go.cr",
        "jcamacho@dinadeco.go.cr"
    ]
}
