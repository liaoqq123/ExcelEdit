from flask import Flask, render_template, request
from user_data import user_data
app = Flask(__name__)


@app.route('/')
def open_login():
    #打开登录页面
    return render_template("Login.html")

@app.route('/login', methods=["POST"])
def login_data():
    #登录并获取用户数据
    username = request.form.get('username')
    password = request.form.get('password')
    for user in user_data:
        if user['username'] == username:
            if user['password'] == password:
                return render_template("Main.html")
            else:
                print('密码错误')
        else:
            print('用户名错误')
    print(username, password)


if __name__ == '__main__':
    app.run()
