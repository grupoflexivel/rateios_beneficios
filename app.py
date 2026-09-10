"""Flask application factory and web entry point."""

from flask import Flask

from rateiosrh.web.routes import web


def create_app() -> Flask:
    """Create the Flask application with the registered web routes."""

    application = Flask(__name__)
    application.register_blueprint(web)
    return application


app = create_app()


if __name__ == "__main__":
    app.run(debug=False)
