from flask import Flask, render_template,redirect
from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField
from wtforms.validators import DataRequired
from uuid import uuid4
from main import get_timetable

app = Flask(__name__)
app.config["SECRET_KEY"] = str(uuid4())


class SubmitForm(FlaskForm):
    sem = StringField("Enter Semester Number (Eg: 2)",default=2, validators=[DataRequired()])
    roll_num = StringField(
        "Enter last three digits of your roll number (Eg: 146)",
        validators=[DataRequired()],
    )
    submit = SubmitField("Submit")


@app.route("/", methods=["GET", "POST"])
def home():
    form = SubmitForm()
    if form.validate_on_submit():
        sem = int(form.sem.data)
        form.sem.data = "2"
        roll_num = int(form.roll_num.data)
        form.roll_num.data = ''
        if roll_num >373 or sem>2 or roll_num < 1 or sem < 1:
            return render_template('error.html',e="Roll number should be between 1 and 373 (both inclusive)",show=False)
        try:
            return redirect(get_timetable(int(roll_num),int(sem)))
        except Exception as e:
            return render_template('error.html',e=e,show=True)

    return render_template("home.html",form=form)


app.run(port=5100, debug=True)
