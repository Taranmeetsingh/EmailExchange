from flask import Flask
from flask_restful import Api, Resource, reqparse
import test_api

app = Flask(__name__)
api = Api(app)


api.add_resource(test_api.User, "/user", "/ta")


if __name__ == '__main__':
    app.run(port = 5005)

