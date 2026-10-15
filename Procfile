release: python manage.py migrate --no-input
web: gunicorn your_project_name.wsgi --bind 0.0.0.0:$PORT --log-file -