# lumber_track/context_processors.py

def user_role(request):
    """Передает роль пользователя в шаблоны"""
    if request.user.is_authenticated:
        # Определяем роль по имени пользователя или группе
        if request.user.username == 'accountant':
            role = 'accountant'
        else:
            role = 'manager'
        return {'user_role': role}
    return {'user_role': None}