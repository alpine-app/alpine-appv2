release: python manage.py migrate --no-input
web: gunicorn AlpineWeb8v3.wsgi --bind 0.0.0.0:$PORT --log-file -