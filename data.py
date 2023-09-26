from flask import Flask, render_template, request, redirect, url_for, flash
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_wtf import FlaskForm
from wtforms import SelectField, DateField, SubmitField
import pandas as pd

import pandas as pd

app = Flask(__name__)
app.secret_key = 'misbah'
def append_to_excel(data, sheet_name):
    try:
        with pd.ExcelFile('master_data.xlsx') as xls:
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
            else:
                df = pd.DataFrame(columns=list(data.keys()))
    except FileNotFoundError:
        df = pd.DataFrame(columns=list(data.keys()))

    df = df.append(data, ignore_index=True)
    
    # Save the data to the Excel file
    with pd.ExcelWriter('master_data.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


@app.route('/')
def index():
    return render_template('dashboard.html')

@app.route('/add_new_course_form')
def add_course_form():
    return render_template('add_new_course.html')

# ... (previous code)

@app.route('/add_new_course', methods=['POST'])
def add_course():
    course_type = request.form['course_type']
    course_name = request.form['course_name']
    duration = request.form['duration']
    online_offline = request.form['online_offline']
    image_url = request.form['image_url']

    course_data = {
        'Course Type': course_type,
        'Course Name': course_name,
        'Duration': duration,
        'Online/Offline': online_offline,
        'Image Url': image_url
    }

    append_to_excel(course_data, 'course details')

    flash('Successfully added a new course', 'success')  # Flash success message for course

    return render_template('add_new_course.html', message='Successfully added a new course')


@app.route('/add_new_batch', methods=['GET', 'POST'])
def add_batch():
    if request.method == 'POST':
        batch_name = request.form['batch_name']
        department = request.form['department']

        batch_data = {
            'Batch Name': batch_name,
            'Department': department,
        }

        append_to_excel(batch_data, 'batch details')

        flash('Batch added successfully', 'success')  # Flash success message

    return render_template('add_new_batch.html')


# Define your Flask-WTF form for allocation
class AllocationForm(FlaskForm):
    batch = SelectField('Batch', choices=[], coerce=str)
    course_type = SelectField('Course Type', choices=[('Foundation', 'Foundation'), ('Induction', 'Induction'), ('Advanced', 'Advanced')], coerce=str)
    course_name = SelectField('Course Name', choices=[], coerce=str)
    start_date = DateField('Start Date', format='%Y-%m-%d')
    end_date = DateField('End Date', format='%Y-%m-%d')
    trainer = SelectField('Trainer', choices=[], coerce=str)
    submit = SubmitField('Allocate')

# Define a function to populate data from Excel files
def fetch_data_from_excel(file_path, sheet_name, column_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if column_name in df.columns:
            return list(df[column_name].dropna().unique())
        else:
            return []
    except FileNotFoundError:
        return []


@app.route('/allocate_course', methods=['GET', 'POST'])
def allocate_course():
    form = AllocationForm()
    
    # Populate batch options dynamically from 'batch details' sheet
    # Update the column names list to include all required columns
    # Update the column names list to include all required columns
    # Update the column names list to include the correct column name
    batch_options = fetch_data_from_excel('master_data.xlsx', 'batch details', 'Batch Name')
    
    form.batch.choices = [(batch, batch) for batch in batch_options]
    
    # Populate course name options dynamically from 'course details' sheet
    course_name_options = fetch_data_from_excel('master_data.xlsx', 'course details', 'Course Name')
    
    form.course_name.choices = [(course, course) for course in course_name_options]
    
    # Populate trainer options dynamically from 'employeedetails' sheet
    trainer_options = fetch_data_from_excel('employeedetails.xlsx', 'Sheet1', 'Employee ID')
   
    form.trainer.choices = [(trainer, trainer) for trainer in trainer_options]
    
    if form.validate_on_submit():
        # Process form submission and allocate the course

        # You can access form data like this:
        batch = form.batch.data
        course_type = form.course_type.data
        course_name = form.course_name.data
        start_date = form.start_date.data
        end_date = form.end_date.data
        trainer = form.trainer.data

        # Add your allocation logic here
        allocation_data = {
            'Batch': batch,
            'Course Type': course_type,
            'Course Name': course_name,
            'Start Date': start_date,
            'End Date': end_date,
            'Trainer': trainer
        }


        # Append the allocation data to 'allocated courses' sheet in 'master_data.xlsx'
        append_to_excel(allocation_data, 'allocated courses')
        

        flash('Course allocated successfully', 'success')  # Flash success message
    
    return render_template('allocate_course.html', form=form,batches=batch_options, courses=course_name_options, trainers=trainer_options)
if __name__ == '__main__':
    app.run(debug=True)
