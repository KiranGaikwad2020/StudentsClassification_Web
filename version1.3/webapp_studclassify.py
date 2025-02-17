from flask import Flask, request, render_template, send_file
import pandas as pd

app = Flask(__name__)

# Function to process file
def categorize_students(input_file, outstanding_min, good_min):
    df = pd.read_excel(input_file)

    def categorize(marks):
        if marks >= outstanding_min:
            return 'Outstanding'
        elif marks >= good_min:
            return 'Good'
        else:
            return 'Poor'

    df['Category'] = df['Marks'].apply(categorize)
    output_file = "categorized_students.xlsx"
    df.to_excel(output_file, index=False)
    return output_file

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        outstanding_min = int(request.form["outstanding"])
        good_min = int(request.form["good"])

        if file:
            file_path = "uploaded.xlsx"
            file.save(file_path)
            output_path = categorize_students(file_path, outstanding_min, good_min)
            return send_file(output_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)

