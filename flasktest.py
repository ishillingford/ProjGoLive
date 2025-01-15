
from flask import Flask, jsonify

app = Flask(__name__)

@app.route('/test-azure')
def hello_world():
    # Return JSON response
    return jsonify({"message": "Hello, World! Your Flask app is running on Azure."})

if __name__ == '__main__':
    app.run(debug=True)
