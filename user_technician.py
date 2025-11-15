from flask import Blueprint, render_template, request, jsonify, send_file
from utils import get_db_connection
import json
import pandas as pd
from io import BytesIO

# Blueprint definition
user_tech_bp = Blueprint('user_tech_bp', __name__, url_prefix='/user_tech')

# Induction Main Page
@user_tech_bp.route('/induction', methods=['GET'])
def induction_list():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT * FROM induction ORDER BY sr_no DESC"
            cursor.execute(sql)
            induction_data = cursor.fetchall()
    except Exception as e:
        print(f"Error fetching induction data: {e}")
        induction_data = []
    finally:
        conn.close()

    return render_template("user/induction.html", induction_data=induction_data)

# Induction Filter Options API
@user_tech_bp.route('/api/induction/filter-options', methods=['GET'])
def get_filter_options():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Get unique values for each filter
            cursor.execute("SELECT DISTINCT plant_location FROM induction WHERE plant_location IS NOT NULL ORDER BY plant_location")
            plants = [row['plant_location'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT batch_number FROM induction WHERE batch_number IS NOT NULL ORDER BY batch_number")
            batches = [row['batch_number'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT learning_hours FROM induction WHERE learning_hours IS NOT NULL ORDER BY learning_hours")
            learning_hours = [str(int(row['learning_hours'])) for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT gender FROM induction WHERE gender IS NOT NULL ORDER BY gender")
            genders = [row['gender'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT employee_category FROM induction WHERE employee_category IS NOT NULL ORDER BY employee_category")
            categories = [row['employee_category'] for row in cursor.fetchall()]
            
            # Get distinct academic years from joined_year
            cursor.execute("SELECT DISTINCT joined_year FROM induction WHERE joined_year IS NOT NULL ORDER BY joined_year")
            academic_years = []
            for row in cursor.fetchall():
                year = row['joined_year']
                # Format as "YYYY/YY"
                formatted_year = f"{int(year)}/{str(int(year)+1)[2:]}"
                academic_years.append(formatted_year)
            
            # Get distinct faculty names
            cursor.execute("SELECT DISTINCT faculty_name FROM induction WHERE faculty_name IS NOT NULL ORDER BY faculty_name")
            faculties = [row['faculty_name'] for row in cursor.fetchall()]
            
            # Get distinct shifts
            cursor.execute("SELECT DISTINCT shift FROM induction WHERE shift IS NOT NULL ORDER BY shift")
            shifts = [row['shift'] for row in cursor.fetchall()]
            
        return jsonify({
            'plants': plants,
            'batches': batches,
            'learning_hours': learning_hours,
            'genders': genders,
            'categories': categories,
            'academic_years': academic_years,
            'faculties': faculties,
            'shifts': shifts
        })
    except Exception as e:
        print(f"Error fetching filter options: {e}")
        return jsonify({
            'plants': [],
            'batches': [],
            'learning_hours': [],
            'genders': [],
            'categories': [],
            'academic_years': [],
            'faculties': [],
            'shifts': []
        }), 500
    finally:
        conn.close()

# Induction Data API
@user_tech_bp.route('/api/induction/data', methods=['GET'])
def get_induction_data():
    try:
        # Get filter parameters from request
        plant = request.args.get('plant', 'all')
        batch = request.args.get('batch', 'all')
        hours = request.args.get('hours', 'all')
        gender = request.args.get('gender', 'all')
        category = request.args.get('category', 'all')
        academic_year = request.args.get('academic_year', 'all')
        faculty = request.args.get('faculty', 'all')
        shift = request.args.get('shift', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        # Pagination parameters
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if plant != 'all':
                where_conditions.append("plant_location = %s")
                params.append(plant)
            
            if batch != 'all':
                where_conditions.append("batch_number = %s")
                params.append(batch)
            
            if hours != 'all':
                where_conditions.append("learning_hours = %s")
                params.append(float(hours))
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if category != 'all':
                where_conditions.append("employee_category = %s")
                params.append(category)
            
            # Handle academic year filter
            if academic_year != 'all':
                # Extract the base year from the formatted academic year (e.g., "2023/24" -> 2023)
                base_year = academic_year.split('/')[0]
                where_conditions.append("joined_year = %s")
                params.append(base_year)
            
            if faculty != 'all':
                where_conditions.append("faculty_name = %s")
                params.append(faculty)
            
            if shift != 'all':
                where_conditions.append("shift = %s")
                params.append(shift)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_from >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_to <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM induction WHERE {where_clause} ORDER BY sr_no DESC LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Get total count for pagination
            count_sql = f"SELECT COUNT(*) as total FROM induction WHERE {where_clause}"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()['total']
            
            # Calculate statistics - Updated to include male and female counts
            stats_sql = f"""
                SELECT 
                    COUNT(DISTINCT batch_number) as batch_count,
                    COUNT(*) as coverage_count,
                    COALESCE(SUM(learning_hours), 0) as learning_hours,
                    COUNT(CASE WHEN gender = 'Male' THEN 1 END) as male_count,
                    COUNT(CASE WHEN gender = 'Female' THEN 1 END) as female_count
                FROM induction 
                WHERE {where_clause}
            """
            cursor.execute(stats_sql, params)
            stats = cursor.fetchone()
            
        # Convert records to list of dictionaries for JSON serialization
        records_list = []
        for record in records:
            # Format academic year for display
            academic_year_display = ""
            if record['joined_year']:
                academic_year_display = f"{int(record['joined_year'])}/{str(int(record['joined_year'])+1)[2:]}"
            
            records_list.append({
                'sr_no': record['sr_no'],
                'ticket_no': record['ticket_no'],
                'name': record['name'],
                'gender': record['gender'],
                'employee_category': record['employee_category'],
                'plant_location': record['plant_location'],
                'academic_year': academic_year_display,
                'date_from': record['date_from'].strftime('%Y-%m-%d') if record['date_from'] else '',
                'date_to': record['date_to'].strftime('%Y-%m-%d') if record['date_to'] else '',
                'shift': record['shift'],
                'learning_hours': int(record['learning_hours']) if record['learning_hours'] else 0,
                'training_name': record['training_name'],
                'batch_number': record['batch_number'],
                'training_venue_name': record['training_venue_name'],
                'faculty_name': record['faculty_name'],
                'subject_name': record['subject_name'],
                'remark': record['remark']
            })
        
        return jsonify({
            'records': records_list,
            'stats': {
                'batch_count': stats['batch_count'] if stats else 0,
                'coverage_count': stats['coverage_count'] if stats else 0,
                'learning_hours': int(stats['learning_hours']) if stats else 0,
                'male_count': stats['male_count'] if stats else 0,
                'female_count': stats['female_count'] if stats else 0
            },
            'total_records': total_count
        })
        
    except Exception as e:
        print(f"Error fetching induction data: {e}")
        return jsonify({
            'records': [],
            'stats': {
                'batch_count': 0,
                'coverage_count': 0,
                'learning_hours': 0,
                'male_count': 0,
                'female_count': 0
            },
            'total_records': 0
        }), 500
    finally:
        conn.close()

# Induction Download Excel API
@user_tech_bp.route('/api/induction/download', methods=['GET'])
def download_induction_data():
    try:
        # Get filter parameters from request
        plant = request.args.get('plant', 'all')
        batch = request.args.get('batch', 'all')
        hours = request.args.get('hours', 'all')
        gender = request.args.get('gender', 'all')
        category = request.args.get('category', 'all')
        academic_year = request.args.get('academic_year', 'all')
        faculty = request.args.get('faculty', 'all')
        shift = request.args.get('shift', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if plant != 'all':
                where_conditions.append("plant_location = %s")
                params.append(plant)
            
            if batch != 'all':
                where_conditions.append("batch_number = %s")
                params.append(batch)
            
            if hours != 'all':
                where_conditions.append("learning_hours = %s")
                params.append(float(hours))
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if category != 'all':
                where_conditions.append("employee_category = %s")
                params.append(category)
            
            # Handle academic year filter
            if academic_year != 'all':
                base_year = academic_year.split('/')[0]
                where_conditions.append("joined_year = %s")
                params.append(base_year)
            
            if faculty != 'all':
                where_conditions.append("faculty_name = %s")
                params.append(faculty)
            
            if shift != 'all':
                where_conditions.append("shift = %s")
                params.append(shift)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_from >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_to <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM induction WHERE {where_clause} ORDER BY sr_no DESC"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Convert records to DataFrame
            df = pd.DataFrame(records)
            
            # Rename columns to proper names
            column_mapping = {
                'sr_no': 'Sr No',
                'ticket_no': 'Ticket No',
                'name': 'Name',
                'gender': 'Gender',
                'employee_category': 'Employee Category',
                'plant_location': 'Plant Location',
                'date_from': 'Date From',
                'date_to': 'Date To',
                'shift': 'Shift',
                'learning_hours': 'Learning Hours',
                'training_name': 'Training Name',
                'batch_number': 'Batch Number',
                'training_venue_name': 'Training Venue Name',
                'faculty_name': 'Faculty Name',
                'subject_name': 'Subject Name',
                'remark': 'Remark'
            }
            
            # Rename the columns
            df = df.rename(columns=column_mapping)
            
            # Create Academic Year column from joined_year and format it
            if 'joined_year' in df.columns:
                df['Academic Year'] = df['joined_year'].apply(
                    lambda x: f"{int(x)}/{str(int(x)+1)[2:]}" if pd.notnull(x) else ''
                )
                # Drop the original joined_year column
                df = df.drop(columns=['joined_year'])
            
            # Format dates
            date_columns = ['Date From', 'Date To']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
            
            # Convert learning_hours to int
            if 'Learning Hours' in df.columns:
                df['Learning Hours'] = df['Learning Hours'].apply(
                    lambda x: int(x) if pd.notnull(x) else 0
                )
            
            # Reorder columns to match the desired order
            desired_columns = [
                'Sr No', 'Ticket No', 'Name', 'Gender', 'Employee Category', 
                'Plant Location', 'Academic Year', 'Date From', 
                'Date To', 'Shift', 'Learning Hours', 'Training Name', 
                'Batch Number', 'Training Venue Name', 'Faculty Name', 
                'Subject Name', 'Remark'
            ]
            
            # Filter to only include columns that exist in the dataframe
            columns_to_use = [col for col in desired_columns if col in df.columns]
            df = df[columns_to_use]
            
            # Create an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Induction Data')
                
                # Adjust column widths for better readability
                worksheet = writer.sheets['Induction Data']
                for idx, col in enumerate(df.columns):
                    # Set column width based on content length
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2  # Add some padding
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)  # Cap width at 50
            
            output.seek(0)
            
            # Return the Excel file as a response
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='induction_data.xlsx'
            )
            
    except Exception as e:
        print(f"Error generating Excel file: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()

# FST Main Page
@user_tech_bp.route('/fst', methods=['GET'])
def fst_list():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT * FROM fst ORDER BY sr_no DESC"
            cursor.execute(sql)
            fst_data = cursor.fetchall()
    except Exception as e:
        print(f"Error fetching FST data: {e}")
        fst_data = []
    finally:
        conn.close()

    return render_template("user/fst.html", fst_data=fst_data)

# FST Filter Options API
@user_tech_bp.route('/api/fst/filter-options', methods=['GET'])
def get_fst_filter_options():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Get unique values for each filter
            cursor.execute("SELECT DISTINCT plant_location FROM fst WHERE plant_location IS NOT NULL ORDER BY plant_location")
            plants = [row['plant_location'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT batch_number FROM fst WHERE batch_number IS NOT NULL ORDER BY batch_number")
            batches = [row['batch_number'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT learning_hours FROM fst WHERE learning_hours IS NOT NULL ORDER BY learning_hours")
            learning_hours = [str(row['learning_hours']) for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT gender FROM fst WHERE gender IS NOT NULL ORDER BY gender")
            genders = [row['gender'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT employee_category FROM fst WHERE employee_category IS NOT NULL ORDER BY employee_category")
            categories = [row['employee_category'] for row in cursor.fetchall()]
            
            # Get distinct academic years from joined_year
            cursor.execute("SELECT DISTINCT joined_year FROM fst WHERE joined_year IS NOT NULL ORDER BY joined_year")
            academic_years = []
            for row in cursor.fetchall():
                year = row['joined_year']
                # Format as "YYYY/YY"
                formatted_year = f"{int(year)}/{str(int(year)+1)[2:]}"
                academic_years.append(formatted_year)
            
            # Get distinct faculty names
            cursor.execute("SELECT DISTINCT faculty_name FROM fst WHERE faculty_name IS NOT NULL ORDER BY faculty_name")
            faculties = [row['faculty_name'] for row in cursor.fetchall()]
            
            # Get distinct shifts
            cursor.execute("SELECT DISTINCT shift FROM fst WHERE shift IS NOT NULL ORDER BY shift")
            shifts = [row['shift'] for row in cursor.fetchall()]
            
            # Get distinct FST cell names
            cursor.execute("SELECT DISTINCT fst_cell_name FROM fst WHERE fst_cell_name IS NOT NULL ORDER BY fst_cell_name")
            fst_cells = [row['fst_cell_name'] for row in cursor.fetchall()]
            
        return jsonify({
            'plants': plants,
            'batches': batches,
            'learning_hours': learning_hours,
            'genders': genders,
            'categories': categories,
            'academic_years': academic_years,
            'faculties': faculties,
            'shifts': shifts,
            'fst_cells': fst_cells
        })
    except Exception as e:
        print(f"Error fetching FST filter options: {e}")
        return jsonify({
            'plants': [],
            'batches': [],
            'learning_hours': [],
            'genders': [],
            'categories': [],
            'academic_years': [],
            'faculties': [],
            'shifts': [],
            'fst_cells': []
        }), 500
    finally:
        conn.close()

# FST Data API
@user_tech_bp.route('/api/fst/data', methods=['GET'])
def get_fst_data():
    try:
        # Get filter parameters from request
        plant = request.args.get('plant', 'all')
        batch = request.args.get('batch', 'all')
        hours = request.args.get('hours', 'all')
        gender = request.args.get('gender', 'all')
        category = request.args.get('category', 'all')
        academic_year = request.args.get('academic_year', 'all')
        faculty = request.args.get('faculty', 'all')
        shift = request.args.get('shift', 'all')
        fst_cell = request.args.get('fst_cell', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        # Pagination parameters
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if plant != 'all':
                where_conditions.append("plant_location = %s")
                params.append(plant)
            
            if batch != 'all':
                where_conditions.append("batch_number = %s")
                params.append(batch)
            
            if hours != 'all':
                where_conditions.append("learning_hours = %s")
                params.append(hours)
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if category != 'all':
                where_conditions.append("employee_category = %s")
                params.append(category)
            
            # Handle academic year filter
            if academic_year != 'all':
                # Extract the base year from the formatted academic year (e.g., "2023/24" -> 2023)
                base_year = academic_year.split('/')[0]
                where_conditions.append("joined_year = %s")
                params.append(base_year)
            
            if faculty != 'all':
                where_conditions.append("faculty_name = %s")
                params.append(faculty)
            
            if shift != 'all':
                where_conditions.append("shift = %s")
                params.append(shift)
            
            if fst_cell != 'all':
                where_conditions.append("fst_cell_name = %s")
                params.append(fst_cell)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_from >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_to <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM fst WHERE {where_clause} ORDER BY sr_no DESC LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Get total count for pagination
            count_sql = f"SELECT COUNT(*) as total FROM fst WHERE {where_clause}"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()['total']
            
            # Calculate statistics - Updated to include male and female counts
            stats_sql = f"""
                SELECT 
                    COUNT(DISTINCT batch_number) as batch_count,
                    COUNT(*) as coverage_count,
                    COALESCE(SUM(CAST(learning_hours AS SIGNED)), 0) as learning_hours,
                    COUNT(CASE WHEN gender = 'Male' THEN 1 END) as male_count,
                    COUNT(CASE WHEN gender = 'Female' THEN 1 END) as female_count
                FROM fst 
                WHERE {where_clause}
            """
            cursor.execute(stats_sql, params)
            stats = cursor.fetchone()
            
        # Convert records to list of dictionaries for JSON serialization
        records_list = []
        for record in records:
            # Format academic year for display
            academic_year_display = ""
            if record['joined_year']:
                academic_year_display = f"{int(record['joined_year'])}/{str(int(record['joined_year'])+1)[2:]}"
            
            records_list.append({
                'sr_no': record['sr_no'],
                'ticket_no': record['ticket_no'],
                'name': record['name'],
                'gender': record['gender'],
                'employee_category': record['employee_category'],
                'plant_location': record['plant_location'],
                'academic_year': academic_year_display,
                'date_from': record['date_from'].strftime('%Y-%m-%d') if record['date_from'] else '',
                'date_to': record['date_to'].strftime('%Y-%m-%d') if record['date_to'] else '',
                'shift': record['shift'],
                'learning_hours': record['learning_hours'],
                'training_name': record['training_name'],
                'batch_number': record['batch_number'],
                'training_venue_name': record['training_venue_name'],
                'faculty_name': record['faculty_name'],
                'fst_cell_name': record['fst_cell_name'],
                'remark': record['remark']
            })
        
        return jsonify({
            'records': records_list,
            'stats': {
                'batch_count': stats['batch_count'] if stats else 0,
                'coverage_count': stats['coverage_count'] if stats else 0,
                'learning_hours': int(stats['learning_hours']) if stats else 0,
                'male_count': stats['male_count'] if stats else 0,
                'female_count': stats['female_count'] if stats else 0
            },
            'total_records': total_count
        })
        
    except Exception as e:
        print(f"Error fetching FST data: {e}")
        return jsonify({
            'records': [],
            'stats': {
                'batch_count': 0,
                'coverage_count': 0,
                'learning_hours': 0,
                'male_count': 0,
                'female_count': 0
            },
            'total_records': 0
        }), 500
    finally:
        conn.close()

# FST Download Excel API
@user_tech_bp.route('/api/fst/download', methods=['GET'])
def download_fst_data():
    try:
        # Get filter parameters from request
        plant = request.args.get('plant', 'all')
        batch = request.args.get('batch', 'all')
        hours = request.args.get('hours', 'all')
        gender = request.args.get('gender', 'all')
        category = request.args.get('category', 'all')
        academic_year = request.args.get('academic_year', 'all')
        faculty = request.args.get('faculty', 'all')
        shift = request.args.get('shift', 'all')
        fst_cell = request.args.get('fst_cell', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if plant != 'all':
                where_conditions.append("plant_location = %s")
                params.append(plant)
            
            if batch != 'all':
                where_conditions.append("batch_number = %s")
                params.append(batch)
            
            if hours != 'all':
                where_conditions.append("learning_hours = %s")
                params.append(hours)
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if category != 'all':
                where_conditions.append("employee_category = %s")
                params.append(category)
            
            # Handle academic year filter
            if academic_year != 'all':
                base_year = academic_year.split('/')[0]
                where_conditions.append("joined_year = %s")
                params.append(base_year)
            
            if faculty != 'all':
                where_conditions.append("faculty_name = %s")
                params.append(faculty)
            
            if shift != 'all':
                where_conditions.append("shift = %s")
                params.append(shift)
            
            if fst_cell != 'all':
                where_conditions.append("fst_cell_name = %s")
                params.append(fst_cell)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_from >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_to <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM fst WHERE {where_clause} ORDER BY sr_no DESC"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Convert records to DataFrame
            df = pd.DataFrame(records)
            
            # Rename columns to proper names
            column_mapping = {
                'sr_no': 'Sr No',
                'ticket_no': 'Ticket No',
                'name': 'Name',
                'gender': 'Gender',
                'employee_category': 'Employee Category',
                'plant_location': 'Plant Location',
                'date_from': 'Date From',
                'date_to': 'Date To',
                'shift': 'Shift',
                'learning_hours': 'Learning Hours',
                'training_name': 'Training Name',
                'batch_number': 'Batch Number',
                'training_venue_name': 'Training Venue Name',
                'faculty_name': 'Faculty Name',
                'fst_cell_name': 'FST Cell Name',
                'remark': 'Remark'
            }
            
            # Rename the columns
            df = df.rename(columns=column_mapping)
            
            # Create Academic Year column from joined_year and format it
            if 'joined_year' in df.columns:
                df['Academic Year'] = df['joined_year'].apply(
                    lambda x: f"{int(x)}/{str(int(x)+1)[2:]}" if pd.notnull(x) else ''
                )
                # Drop the original joined_year column
                df = df.drop(columns=['joined_year'])
            
            # Format dates
            date_columns = ['Date From', 'Date To']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
            
            # Reorder columns to match the desired order
            desired_columns = [
                'Sr No', 'Ticket No', 'Name', 'Gender', 'Employee Category', 
                'Plant Location', 'Academic Year', 'Date From', 
                'Date To', 'Shift', 'Learning Hours', 'Training Name', 
                'Batch Number', 'Training Venue Name', 'Faculty Name', 
                'FST Cell Name', 'Remark'
            ]
            
            # Filter to only include columns that exist in the dataframe
            columns_to_use = [col for col in desired_columns if col in df.columns]
            df = df[columns_to_use]
            
            # Create an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='FST Data')
                
                # Adjust column widths for better readability
                worksheet = writer.sheets['FST Data']
                for idx, col in enumerate(df.columns):
                    # Set column width based on content length
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2  # Add some padding
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)  # Cap width at 50
            
            output.seek(0)
            
            # Return the Excel file as a response
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='fst_data.xlsx'
            )
            
    except Exception as e:
        print(f"Error generating FST Excel file: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()
# Pragati Main Page
@user_tech_bp.route('/pragati', methods=['GET'])
def pragati_list():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT * FROM pragati ORDER BY sr_no DESC"
            cursor.execute(sql)
            pragati_data = cursor.fetchall()
    except Exception as e:
        print(f"Error fetching Pragati data: {e}")
        pragati_data = []
    finally:
        conn.close()

    return render_template("user/pragati.html", pragati_data=pragati_data)

# Pragati Filter Options API
@user_tech_bp.route('/api/pragati/filter-options', methods=['GET'])
def get_pragati_filter_options():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Get unique values for each filter
            cursor.execute("SELECT DISTINCT factory FROM pragati WHERE factory IS NOT NULL ORDER BY factory")
            factories = [row['factory'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT pragati_batch_number FROM pragati WHERE pragati_batch_number IS NOT NULL ORDER BY pragati_batch_number")
            batches = [row['pragati_batch_number'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT diploma_name FROM pragati WHERE diploma_name IS NOT NULL ORDER BY diploma_name")
            diplomas = [row['diploma_name'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT gender FROM pragati WHERE gender IS NOT NULL ORDER BY gender")
            genders = [row['gender'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT employee_category FROM pragati WHERE employee_category IS NOT NULL ORDER BY employee_category")
            categories = [row['employee_category'] for row in cursor.fetchall()]
            
            # Get distinct course joining years
            cursor.execute("SELECT DISTINCT course_joining_year FROM pragati WHERE course_joining_year IS NOT NULL ORDER BY course_joining_year")
            course_years = [row['course_joining_year'] for row in cursor.fetchall()]
            
            # Get distinct final results
            cursor.execute("SELECT DISTINCT final_result FROM pragati WHERE final_result IS NOT NULL ORDER BY final_result")
            final_results = [row['final_result'] for row in cursor.fetchall()]
            
            # Get distinct training names
            cursor.execute("SELECT DISTINCT training_name FROM pragati WHERE training_name IS NOT NULL ORDER BY training_name")
            trainings = [row['training_name'] for row in cursor.fetchall()]
            
        return jsonify({
            'factories': factories,
            'batches': batches,
            'diplomas': diplomas,
            'genders': genders,
            'categories': categories,
            'course_years': course_years,
            'final_results': final_results,
            'trainings': trainings
        })
    except Exception as e:
        print(f"Error fetching Pragati filter options: {e}")
        return jsonify({
            'factories': [],
            'batches': [],
            'diplomas': [],
            'genders': [],
            'categories': [],
            'course_years': [],
            'final_results': [],
            'trainings': []
        }), 500
    finally:
        conn.close()

# Pragati Data API
@user_tech_bp.route('/api/pragati/data', methods=['GET'])
def get_pragati_data():
    try:
        # Get filter parameters from request
        factory = request.args.get('factory', 'all')
        batch = request.args.get('batch', 'all')
        diploma = request.args.get('diploma', 'all')
        gender = request.args.get('gender', 'all')
        category = request.args.get('category', 'all')
        course_year = request.args.get('course_year', 'all')
        final_result = request.args.get('final_result', 'all')
        training = request.args.get('training', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        tab = request.args.get('tab', 'overall')
        
        # Pagination parameters
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if factory != 'all':
                where_conditions.append("factory = %s")
                params.append(factory)
            
            if batch != 'all':
                where_conditions.append("pragati_batch_number = %s")
                params.append(batch)
            
            if diploma != 'all':
                where_conditions.append("diploma_name = %s")
                params.append(diploma)
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if category != 'all':
                where_conditions.append("employee_category = %s")
                params.append(category)
            
            if course_year != 'all':
                where_conditions.append("course_joining_year = %s")
                params.append(course_year)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training != 'all':
                where_conditions.append("training_name = %s")
                params.append(training)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_joining <= %s")
                params.append(end_date)
            
            # Add tab-specific conditions based on result columns
            if tab == 'live':
                # Students whose final result is not yet declared
                where_conditions.append("(final_result IS NULL OR final_result = '')")
            elif tab == 'first-year':
                # Students in first year: all three result columns are empty
                where_conditions.append("""
                    (first_year_result IS NULL OR first_year_result = '') 
                    AND (second_year_result IS NULL OR second_year_result = '') 
                    AND (final_result IS NULL OR final_result = '')
                """)
            elif tab == 'second-year':
                # Students in second year: first year result filled, second and final empty
                where_conditions.append("""
                    (first_year_result IS NOT NULL AND first_year_result != '') 
                    AND (second_year_result IS NULL OR second_year_result = '') 
                    AND (final_result IS NULL OR final_result = '')
                """)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM pragati WHERE {where_clause} ORDER BY sr_no DESC LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Get total count for pagination
            count_sql = f"SELECT COUNT(*) as total FROM pragati WHERE {where_clause}"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()['total']
            
            # Calculate statistics
            stats_sql = f"""
                SELECT 
                    COUNT(DISTINCT pragati_batch_number) as batch_count,
                    COUNT(*) as coverage_count,
                    COUNT(CASE WHEN gender = 'Male' THEN 1 END) as male_count,
                    COUNT(CASE WHEN gender = 'Female' THEN 1 END) as female_count,
                    COUNT(CASE WHEN final_result = 'Pass' THEN 1 END) as pass_count,
                    COUNT(CASE WHEN final_result = 'Fail' THEN 1 END) as fail_count
                FROM pragati 
                WHERE {where_clause}
            """
            cursor.execute(stats_sql, params)
            stats = cursor.fetchone()
            
        # Convert records to list of dictionaries for JSON serialization
        records_list = []
        for record in records:
            records_list.append({
                'sr_no': record['sr_no'],
                'ticket_no': record['ticket_no'],
                'name': record['name'],
                'gender': record['gender'],
                'employee_category': record['employee_category'],
                'factory': record['factory'],
                'course_joining_year': record['course_joining_year'],
                'date_of_joining': record['date_of_joining'].strftime('%Y-%m-%d') if record['date_of_joining'] else '',
                'pragati_batch_number': record['pragati_batch_number'],
                'diploma_name': record['diploma_name'],
                'first_year_result': record['first_year_result'],
                'second_year_result': record['second_year_result'],
                'final_result': record['final_result'],
                'training_name': record['training_name'],
                'remark': record['remark']
            })
        
        return jsonify({
            'records': records_list,
            'stats': {
                'batch_count': stats['batch_count'] if stats else 0,
                'coverage_count': stats['coverage_count'] if stats else 0,
                'male_count': stats['male_count'] if stats else 0,
                'female_count': stats['female_count'] if stats else 0,
                'pass_count': stats['pass_count'] if stats else 0,
                'fail_count': stats['fail_count'] if stats else 0
            },
            'total_records': total_count
        })
        
    except Exception as e:
        print(f"Error fetching Pragati data: {e}")
        return jsonify({
            'records': [],
            'stats': {
                'batch_count': 0,
                'coverage_count': 0,
                'male_count': 0,
                'female_count': 0,
                'pass_count': 0,
                'fail_count': 0
            },
            'total_records': 0
        }), 500
    finally:
        conn.close()

# Pragati Download Excel API
@user_tech_bp.route('/api/pragati/download', methods=['GET'])
def download_pragati_data():
    try:
        # Get filter parameters from request
        factory = request.args.get('factory', 'all')
        batch = request.args.get('batch', 'all')
        diploma = request.args.get('diploma', 'all')
        gender = request.args.get('gender', 'all')
        category = request.args.get('category', 'all')
        course_year = request.args.get('course_year', 'all')
        final_result = request.args.get('final_result', 'all')
        training = request.args.get('training', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        tab = request.args.get('tab', 'overall')
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if factory != 'all':
                where_conditions.append("factory = %s")
                params.append(factory)
            
            if batch != 'all':
                where_conditions.append("pragati_batch_number = %s")
                params.append(batch)
            
            if diploma != 'all':
                where_conditions.append("diploma_name = %s")
                params.append(diploma)
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if category != 'all':
                where_conditions.append("employee_category = %s")
                params.append(category)
            
            if course_year != 'all':
                where_conditions.append("course_joining_year = %s")
                params.append(course_year)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training != 'all':
                where_conditions.append("training_name = %s")
                params.append(training)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_joining <= %s")
                params.append(end_date)
            
            # Add tab-specific conditions based on result columns
            if tab == 'live':
                # Students whose final result is not yet declared
                where_conditions.append("(final_result IS NULL OR final_result = '')")
            elif tab == 'first-year':
                # Students in first year: all three result columns are empty
                where_conditions.append("""
                    (first_year_result IS NULL OR first_year_result = '') 
                    AND (second_year_result IS NULL OR second_year_result = '') 
                    AND (final_result IS NULL OR final_result = '')
                """)
            elif tab == 'second-year':
                # Students in second year: first year result filled, second and final empty
                where_conditions.append("""
                    (first_year_result IS NOT NULL AND first_year_result != '') 
                    AND (second_year_result IS NULL OR second_year_result = '') 
                    AND (final_result IS NULL OR final_result = '')
                """)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM pragati WHERE {where_clause} ORDER BY sr_no DESC"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Convert records to DataFrame
            df = pd.DataFrame(records)
            
            # Rename columns to proper names
            column_mapping = {
                'sr_no': 'Sr No',
                'ticket_no': 'Ticket No',
                'name': 'Name',
                'gender': 'Gender',
                'employee_category': 'Employee Category',
                'factory': 'Factory',
                'course_joining_year': 'Course Joining Year',
                'date_of_joining': 'Date of Joining',
                'pragati_batch_number': 'Pragati Batch No',
                'diploma_name': 'Diploma Name',
                'first_year_result': '1st Year Result',
                'second_year_result': '2nd Year Result',
                'final_result': 'Final Result',
                'training_name': 'Training Name',
                'remark': 'Remark'
            }
            
            # Rename the columns
            df = df.rename(columns=column_mapping)
            
            # Format dates
            date_columns = ['Date of Joining']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
            
            # Reorder columns to match the desired order
            desired_columns = [
                'Sr No', 'Ticket No', 'Name', 'Gender', 'Employee Category', 
                'Factory', 'Course Joining Year', 'Date of Joining', 
                'Pragati Batch No', 'Diploma Name', '1st Year Result', 
                '2nd Year Result', 'Final Result', 'Training Name', 'Remark'
            ]
            
            # Filter to only include columns that exist in the dataframe
            columns_to_use = [col for col in desired_columns if col in df.columns]
            df = df[columns_to_use]
            
            # Create an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Pragati Data')
                
                # Adjust column widths for better readability
                worksheet = writer.sheets['Pragati Data']
                for idx, col in enumerate(df.columns):
                    # Set column width based on content length
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2  # Add some padding
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)  # Cap width at 50
            
            output.seek(0)
            
            # Return the Excel file as a response
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='pragati_data.xlsx'
            )
            
    except Exception as e:
        print(f"Error generating Pragati Excel file: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()

import datetime

# FTA Main Page
@user_tech_bp.route('/fta', methods=['GET'])
def fta_list():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT * FROM fta ORDER BY sr_no DESC"
            cursor.execute(sql)
            fta_data = cursor.fetchall()
    except Exception as e:
        print(f"Error fetching FTA data: {e}")
        fta_data = []
    finally:
        conn.close()

    return render_template("user/fta.html", fta_data=fta_data)

# FTA Filter Options API
@user_tech_bp.route('/api/fta/filter-options', methods=['GET'])
def get_fta_filter_options():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Get unique values for each filter
            cursor.execute("SELECT DISTINCT gender FROM fta WHERE gender IS NOT NULL ORDER BY gender")
            genders = [row['gender'] for row in cursor.fetchall()]
            
            # Get distinct academic years from joining_year
            cursor.execute("SELECT DISTINCT joining_year FROM fta WHERE joining_year IS NOT NULL ORDER BY joining_year")
            academic_years = []
            for row in cursor.fetchall():
                year = row['joining_year']
                # Format as "YYYY/YY"
                formatted_year = f"{int(year)}/{str(int(year)+1)[2:]}"
                academic_years.append(formatted_year)
            
            # Get distinct faculty names
            cursor.execute("SELECT DISTINCT faculty_name FROM fta WHERE faculty_name IS NOT NULL ORDER BY faculty_name")
            faculties = [row['faculty_name'] for row in cursor.fetchall()]
            
            # Get distinct FTA batch numbers
            cursor.execute("SELECT DISTINCT fta_batch_number FROM fta WHERE fta_batch_number IS NOT NULL ORDER BY fta_batch_number")
            fta_batches = [row['fta_batch_number'] for row in cursor.fetchall()]
            
            # Get distinct trades
            cursor.execute("SELECT DISTINCT trade FROM fta WHERE trade IS NOT NULL ORDER BY trade")
            trades = [row['trade'] for row in cursor.fetchall()]
            
            # Get distinct all_women_batch values
            cursor.execute("SELECT DISTINCT all_women_batch FROM fta WHERE all_women_batch IS NOT NULL ORDER BY all_women_batch")
            all_women_batches = [row['all_women_batch'] for row in cursor.fetchall()]
            
            # Get distinct second_year_inplant_shop values
            cursor.execute("SELECT DISTINCT second_year_inplant_shop FROM fta WHERE second_year_inplant_shop IS NOT NULL ORDER BY second_year_inplant_shop")
            second_year_inplants = [row['second_year_inplant_shop'] for row in cursor.fetchall()]
            
            # Get distinct final_result values
            cursor.execute("SELECT DISTINCT final_result FROM fta WHERE final_result IS NOT NULL ORDER BY final_result")
            final_results = [row['final_result'] for row in cursor.fetchall()]
            
            # Get distinct training names
            cursor.execute("SELECT DISTINCT training_name FROM fta WHERE training_name IS NOT NULL ORDER BY training_name")
            training_names = [row['training_name'] for row in cursor.fetchall()]
            
        return jsonify({
            'genders': genders,
            'academic_years': academic_years,
            'faculties': faculties,
            'fta_batches': fta_batches,
            'trades': trades,
            'all_women_batches': all_women_batches,
            'second_year_inplants': second_year_inplants,
            'final_results': final_results,
            'training_names': training_names
        })
    except Exception as e:
        print(f"Error fetching FTA filter options: {e}")
        return jsonify({
            'genders': [],
            'academic_years': [],
            'faculties': [],
            'fta_batches': [],
            'trades': [],
            'all_women_batches': [],
            'second_year_inplants': [],
            'final_results': [],
            'training_names': []
        }), 500
    finally:
        conn.close()

# FTA Data API
@user_tech_bp.route('/api/fta/data', methods=['GET'])
def get_fta_data():
    try:
        # Get filter parameters from request
        gender = request.args.get('gender', 'all')
        academic_year = request.args.get('academic_year', 'all')
        faculty = request.args.get('faculty', 'all')
        fta_batch = request.args.get('fta_batch', 'all')
        trade = request.args.get('trade', 'all')
        all_women_batch = request.args.get('all_women_batch', 'all')
        second_year_inplant = request.args.get('second_year_inplant', 'all')
        final_result = request.args.get('final_result', 'all')
        training_name = request.args.get('training_name', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        # Pagination parameters
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            # Handle academic year filter
            if academic_year != 'all':
                # Extract the base year from the formatted academic year (e.g., "2023/24" -> 2023)
                base_year = academic_year.split('/')[0]
                where_conditions.append("joining_year = %s")
                params.append(base_year)
            
            if faculty != 'all':
                where_conditions.append("faculty_name = %s")
                params.append(faculty)
            
            if fta_batch != 'all':
                where_conditions.append("fta_batch_number = %s")
                params.append(fta_batch)
            
            if trade != 'all':
                where_conditions.append("trade = %s")
                params.append(trade)
            
            if all_women_batch != 'all':
                where_conditions.append("all_women_batch = %s")
                params.append(all_women_batch)
            
            if second_year_inplant != 'all':
                where_conditions.append("second_year_inplant_shop = %s")
                params.append(second_year_inplant)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training_name != 'all':
                where_conditions.append("training_name = %s")
                params.append(training_name)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_separation <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM fta WHERE {where_clause} ORDER BY sr_no DESC LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Get total count for pagination
            count_sql = f"SELECT COUNT(*) as total FROM fta WHERE {where_clause}"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()['total']
            
            # Calculate statistics - Updated to include male and female counts
            stats_sql = f"""
                SELECT 
                    COUNT(DISTINCT fta_batch_number) as batch_count,
                    COUNT(*) as coverage_count,
                    COUNT(CASE WHEN all_women_batch = 'Yes' THEN 1 END) as women_batch_count,
                    COUNT(CASE WHEN second_year_inplant_shop = 'Yes' THEN 1 END) as inplant_shop_count,
                    COUNT(CASE WHEN gender = 'Male' THEN 1 END) as male_count,
                    COUNT(CASE WHEN gender = 'Female' THEN 1 END) as female_count
                FROM fta 
                WHERE {where_clause}
            """
            cursor.execute(stats_sql, params)
            stats = cursor.fetchone()
            
        # Convert records to list of dictionaries for JSON serialization
        records_list = []
        for record in records:
            # Format academic year for display
            academic_year_display = ""
            if record['joining_year']:
                academic_year_display = f"{int(record['joining_year'])}/{str(int(record['joining_year'])+1)[2:]}"
            
            records_list.append({
                'sr_no': record['sr_no'],
                'ticket_no': record['ticket_no'],
                'name': record['name'],
                'gender': record['gender'],
                'academic_year': academic_year_display,
                'date_of_joining': record['date_of_joining'].strftime('%Y-%m-%d') if record['date_of_joining'] else '',
                'fta_batch_number': record['fta_batch_number'],
                'date_of_separation': record['date_of_separation'].strftime('%Y-%m-%d') if record['date_of_separation'] else '',
                'trade': record['trade'],
                'all_women_batch': record['all_women_batch'],
                'second_year_inplant_shop': record['second_year_inplant_shop'],
                'faculty_name': record['faculty_name'],
                'final_result': record['final_result'],
                'training_name': record['training_name']
            })
        
        return jsonify({
            'records': records_list,
            'stats': {
                'batch_count': stats['batch_count'] if stats else 0,
                'coverage_count': stats['coverage_count'] if stats else 0,
                'women_batch_count': stats['women_batch_count'] if stats else 0,
                'inplant_shop_count': stats['inplant_shop_count'] if stats else 0,
                'male_count': stats['male_count'] if stats else 0,
                'female_count': stats['female_count'] if stats else 0
            },
            'total_records': total_count
        })
        
    except Exception as e:
        print(f"Error fetching FTA data: {e}")
        return jsonify({
            'records': [],
            'stats': {
                'batch_count': 0,
                'coverage_count': 0,
                'women_batch_count': 0,
                'inplant_shop_count': 0,
                'male_count': 0,
                'female_count': 0
            },
            'total_records': 0
        }), 500
    finally:
        conn.close()

# FTA Download Excel API
@user_tech_bp.route('/api/fta/download', methods=['GET'])
def download_fta_data():
    try:
        # Get filter parameters from request
        gender = request.args.get('gender', 'all')
        academic_year = request.args.get('academic_year', 'all')
        faculty = request.args.get('faculty', 'all')
        fta_batch = request.args.get('fta_batch', 'all')
        trade = request.args.get('trade', 'all')
        all_women_batch = request.args.get('all_women_batch', 'all')
        second_year_inplant = request.args.get('second_year_inplant', 'all')
        final_result = request.args.get('final_result', 'all')
        training_name = request.args.get('training_name', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            # Handle academic year filter
            if academic_year != 'all':
                base_year = academic_year.split('/')[0]
                where_conditions.append("joining_year = %s")
                params.append(base_year)
            
            if faculty != 'all':
                where_conditions.append("faculty_name = %s")
                params.append(faculty)
            
            if fta_batch != 'all':
                where_conditions.append("fta_batch_number = %s")
                params.append(fta_batch)
            
            if trade != 'all':
                where_conditions.append("trade = %s")
                params.append(trade)
            
            if all_women_batch != 'all':
                where_conditions.append("all_women_batch = %s")
                params.append(all_women_batch)
            
            if second_year_inplant != 'all':
                where_conditions.append("second_year_inplant_shop = %s")
                params.append(second_year_inplant)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training_name != 'all':
                where_conditions.append("training_name = %s")
                params.append(training_name)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_separation <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM fta WHERE {where_clause} ORDER BY sr_no DESC"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Convert records to DataFrame
            df = pd.DataFrame(records)
            
            # Rename columns to proper names
            column_mapping = {
                'sr_no': 'Sr No',
                'ticket_no': 'Ticket No',
                'name': 'Name',
                'gender': 'Gender',
                'joining_year': 'Joining Year',
                'date_of_joining': 'Date of Joining',
                'fta_batch_number': 'FTA Batch Number',
                'date_of_separation': 'Date of Separation',
                'trade': 'Trade',
                'all_women_batch': 'All Women Batch',
                'second_year_inplant_shop': 'Second Year Implant Shop',
                'faculty_name': 'Faculty Name',
                'final_result': 'Final Result',
                'training_name': 'Training Name'
            }
            
            # Rename the columns
            df = df.rename(columns=column_mapping)
            
            # Create Academic Year column from joining_year and format it
            if 'Joining Year' in df.columns:
                df['Academic Year'] = df['Joining Year'].apply(
                    lambda x: f"{int(x)}/{str(int(x)+1)[2:]}" if pd.notnull(x) else ''
                )
                # Drop the original joining_year column
                df = df.drop(columns=['Joining Year'])
            
            # Format dates
            date_columns = ['Date of Joining', 'Date of Separation']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
            
            # Reorder columns to match the desired order
            desired_columns = [
                'Sr No', 'Ticket No', 'Name', 'Gender', 'Academic Year', 
                'Date of Joining', 'FTA Batch Number', 'Date of Separation', 
                'Trade', 'All Women Batch', 'Second Year Implant Shop', 
                'Faculty Name', 'Final Result', 'Training Name'
            ]
            
            # Filter to only include columns that exist in the dataframe
            columns_to_use = [col for col in desired_columns if col in df.columns]
            df = df[columns_to_use]
            
            # Create an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='FTA Data')
                
                # Adjust column widths for better readability
                worksheet = writer.sheets['FTA Data']
                for idx, col in enumerate(df.columns):
                    # Set column width based on content length
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2  # Add some padding
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)  # Cap width at 50
            
            output.seek(0)
            
            # Return the Excel file as a response
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='fta_data.xlsx'
            )
            
    except Exception as e:
        print(f"Error generating FTA Excel file: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()

# JTA Main Page
@user_tech_bp.route('/jta', methods=['GET'])
def jta_list():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT * FROM jta ORDER BY sr_no DESC"
            cursor.execute(sql)
            jta_data = cursor.fetchall()
    except Exception as e:
        print(f"Error fetching JTA data: {e}")
        jta_data = []
    finally:
        conn.close()

    return render_template("user/jta.html", jta_data=jta_data)

# JTA Filter Options API
@user_tech_bp.route('/api/jta/filter-options', methods=['GET'])
def get_jta_filter_options():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Get unique values for each filter
            cursor.execute("SELECT DISTINCT gender FROM jta WHERE gender IS NOT NULL ORDER BY gender")
            genders = [row['gender'] for row in cursor.fetchall()]
            
            # Get distinct academic years from joining_year
            cursor.execute("SELECT DISTINCT joining_year FROM jta WHERE joining_year IS NOT NULL ORDER BY joining_year")
            academic_years = []
            for row in cursor.fetchall():
                year = row['joining_year']
                # Format as "YYYY/YY"
                formatted_year = f"{int(year)}/{str(int(year)+1)[2:]}"
                academic_years.append(formatted_year)
            
            # Get distinct JTA batch numbers
            cursor.execute("SELECT DISTINCT jta_batch_number FROM jta WHERE jta_batch_number IS NOT NULL ORDER BY jta_batch_number")
            jta_batches = [row['jta_batch_number'] for row in cursor.fetchall()]
            
            # Get distinct trades
            cursor.execute("SELECT DISTINCT trade FROM jta WHERE trade IS NOT NULL ORDER BY trade")
            trades = [row['trade'] for row in cursor.fetchall()]
            
            # Get distinct final_result values
            cursor.execute("SELECT DISTINCT final_result FROM jta WHERE final_result IS NOT NULL ORDER BY final_result")
            final_results = [row['final_result'] for row in cursor.fetchall()]
            
            # Get distinct training names
            cursor.execute("SELECT DISTINCT training_name FROM jta WHERE training_name IS NOT NULL ORDER BY training_name")
            training_names = [row['training_name'] for row in cursor.fetchall()]
            
            # Get distinct employee categories
            cursor.execute("SELECT DISTINCT employee_category FROM jta WHERE employee_category IS NOT NULL ORDER BY employee_category")
            employee_categories = [row['employee_category'] for row in cursor.fetchall()]
            
            # Get distinct status values
            cursor.execute("SELECT DISTINCT status FROM jta WHERE status IS NOT NULL ORDER BY status")
            statuses = [row['status'] for row in cursor.fetchall()]
            
        return jsonify({
            'genders': genders,
            'academic_years': academic_years,
            'jta_batches': jta_batches,
            'trades': trades,
            'final_results': final_results,
            'training_names': training_names,
            'employee_categories': employee_categories,
            'statuses': statuses
        })
    except Exception as e:
        print(f"Error fetching JTA filter options: {e}")
        return jsonify({
            'genders': [],
            'academic_years': [],
            'jta_batches': [],
            'trades': [],
            'final_results': [],
            'training_names': [],
            'employee_categories': [],
            'statuses': []
        }), 500
    finally:
        conn.close()

# JTA Data API
@user_tech_bp.route('/api/jta/data', methods=['GET'])
def get_jta_data():
    try:
        # Get filter parameters from request
        gender = request.args.get('gender', 'all')
        academic_year = request.args.get('academic_year', 'all')
        jta_batch = request.args.get('jta_batch', 'all')
        trade = request.args.get('trade', 'all')
        final_result = request.args.get('final_result', 'all')
        training_name = request.args.get('training_name', 'all')
        employee_category = request.args.get('employee_category', 'all')
        status = request.args.get('status', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        # Pagination parameters
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            # Handle academic year filter
            if academic_year != 'all':
                # Extract the base year from the formatted academic year (e.g., "2023/24" -> 2023)
                base_year = academic_year.split('/')[0]
                where_conditions.append("joining_year = %s")
                params.append(base_year)
            
            if jta_batch != 'all':
                where_conditions.append("jta_batch_number = %s")
                params.append(jta_batch)
            
            if trade != 'all':
                where_conditions.append("trade = %s")
                params.append(trade)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training_name != 'all':
                where_conditions.append("training_name = %s")
                params.append(training_name)
            
            if employee_category != 'all':
                where_conditions.append("employee_category = %s")
                params.append(employee_category)
            
            if status != 'all':
                where_conditions.append("status = %s")
                params.append(status)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_separation <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM jta WHERE {where_clause} ORDER BY sr_no DESC LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Get total count for pagination
            count_sql = f"SELECT COUNT(*) as total FROM jta WHERE {where_clause}"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()['total']
            
            # Calculate statistics with additional fields
            stats_sql = f"""
                SELECT 
                    COUNT(DISTINCT jta_batch_number) as batch_count,
                    COUNT(*) as coverage_count,
                    COUNT(CASE WHEN gender = 'Male' THEN 1 END) as male_count,
                    COUNT(CASE WHEN gender = 'Female' THEN 1 END) as female_count,
                    COUNT(CASE WHEN status = 'Active' THEN 1 END) as live_count,
                    COUNT(CASE WHEN status IN ('Separated', 'Terminated', 'Resigned', 'Retired') THEN 1 END) as separated_count
                FROM jta 
                WHERE {where_clause}
            """
            cursor.execute(stats_sql, params)
            stats = cursor.fetchone()
            
        # Convert records to list of dictionaries for JSON serialization
        records_list = []
        for record in records:
            # Format academic year for display
            academic_year_display = ""
            if record['joining_year']:
                academic_year_display = f"{int(record['joining_year'])}/{str(int(record['joining_year'])+1)[2:]}"
            
            records_list.append({
                'sr_no': record['sr_no'],
                'ticket_no': record['ticket_no'],
                'name': record['name'],
                'gender': record['gender'],
                'academic_year': academic_year_display,
                'date_of_joining': record['date_of_joining'].strftime('%Y-%m-%d') if record['date_of_joining'] else '',
                'jta_batch_number': record['jta_batch_number'],
                'date_of_separation': record['date_of_separation'].strftime('%Y-%m-%d') if record['date_of_separation'] else '',
                'trade': record['trade'],
                'final_result': record['final_result'],
                'training_name': record['training_name'],
                'employee_category': record['employee_category'],
                'status': record['status']
            })
        
        return jsonify({
            'records': records_list,
            'stats': {
                'batch_count': stats['batch_count'] if stats else 0,
                'coverage_count': stats['coverage_count'] if stats else 0,
                'male_count': stats['male_count'] if stats else 0,
                'female_count': stats['female_count'] if stats else 0,
                'live_count': stats['live_count'] if stats else 0,
                'separated_count': stats['separated_count'] if stats else 0
            },
            'total_records': total_count
        })
        
    except Exception as e:
        print(f"Error fetching JTA data: {e}")
        return jsonify({
            'records': [],
            'stats': {
                'batch_count': 0,
                'coverage_count': 0,
                'male_count': 0,
                'female_count': 0,
                'live_count': 0,
                'separated_count': 0
            },
            'total_records': 0
        }), 500
    finally:
        conn.close()

# JTA Download Excel API
@user_tech_bp.route('/api/jta/download', methods=['GET'])
def download_jta_data():
    try:
        # Get filter parameters from request
        gender = request.args.get('gender', 'all')
        academic_year = request.args.get('academic_year', 'all')
        jta_batch = request.args.get('jta_batch', 'all')
        trade = request.args.get('trade', 'all')
        final_result = request.args.get('final_result', 'all')
        training_name = request.args.get('training_name', 'all')
        employee_category = request.args.get('employee_category', 'all')
        status = request.args.get('status', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            # Handle academic year filter
            if academic_year != 'all':
                base_year = academic_year.split('/')[0]
                where_conditions.append("joining_year = %s")
                params.append(base_year)
            
            if jta_batch != 'all':
                where_conditions.append("jta_batch_number = %s")
                params.append(jta_batch)
            
            if trade != 'all':
                where_conditions.append("trade = %s")
                params.append(trade)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training_name != 'all':
                where_conditions.append("training_name = %s")
                params.append(training_name)
            
            if employee_category != 'all':
                where_conditions.append("employee_category = %s")
                params.append(employee_category)
            
            if status != 'all':
                where_conditions.append("status = %s")
                params.append(status)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_separation <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM jta WHERE {where_clause} ORDER BY sr_no DESC"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Convert records to DataFrame
            df = pd.DataFrame(records)
            
            # Rename columns to proper names
            column_mapping = {
                'sr_no': 'Sr No',
                'ticket_no': 'Ticket No',
                'name': 'Name',
                'gender': 'Gender',
                'joining_year': 'Joining Year',
                'date_of_joining': 'Date of Joining',
                'jta_batch_number': 'JTA Batch Number',
                'date_of_separation': 'Date of Separation',
                'trade': 'Trade',
                'final_result': 'Final Result',
                'training_name': 'Training Name',
                'employee_category': 'Employee Category',
                'status': 'Status'
            }
            
            # Rename the columns
            df = df.rename(columns=column_mapping)
            
            # Create Academic Year column from joining_year and format it
            if 'Joining Year' in df.columns:
                df['Academic Year'] = df['Joining Year'].apply(
                    lambda x: f"{int(x)}/{str(int(x)+1)[2:]}" if pd.notnull(x) else ''
                )
                # Drop the original joining_year column
                df = df.drop(columns=['Joining Year'])
            
            # Format dates
            date_columns = ['Date of Joining', 'Date of Separation']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
            
            # Reorder columns to match the desired order
            desired_columns = [
                'Sr No', 'Ticket No', 'Name', 'Gender', 'Academic Year', 
                'Date of Joining', 'JTA Batch Number', 'Date of Separation', 
                'Trade', 'Final Result', 'Training Name', 'Employee Category', 'Status'
            ]
            
            # Filter to only include columns that exist in the dataframe
            columns_to_use = [col for col in desired_columns if col in df.columns]
            df = df[columns_to_use]
            
            # Create an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='JTA Data')
                
                # Adjust column widths for better readability
                worksheet = writer.sheets['JTA Data']
                for idx, col in enumerate(df.columns):
                    # Set column width based on content length
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2  # Add some padding
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)  # Cap width at 50
            
            output.seek(0)
            
            # Return the Excel file as a response
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='jta_data.xlsx'
            )
            
    except Exception as e:
        print(f"Error generating JTA Excel file: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()
# TA Main Page
@user_tech_bp.route('/ta', methods=['GET'])
def ta_list():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT * FROM ta ORDER BY sr_no DESC"
            cursor.execute(sql)
            ta_data = cursor.fetchall()
    except Exception as e:
        print(f"Error fetching TA data: {e}")
        ta_data = []
    finally:
        conn.close()

    return render_template("user/ta.html", ta_data=ta_data)

# TA Filter Options API
@user_tech_bp.route('/api/ta/filter-options', methods=['GET'])
def get_ta_filter_options():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Get unique values for each filter
            cursor.execute("SELECT DISTINCT gender FROM ta WHERE gender IS NOT NULL ORDER BY gender")
            genders = [row['gender'] for row in cursor.fetchall()]
            
            # Get distinct academic years from joining_year
            cursor.execute("SELECT DISTINCT joining_year FROM ta WHERE joining_year IS NOT NULL ORDER BY joining_year")
            academic_years = []
            for row in cursor.fetchall():
                year = row['joining_year']
                # Format as "YYYY/YY"
                formatted_year = f"{int(year)}/{str(int(year)+1)[2:]}"
                academic_years.append(formatted_year)
            
            # Get distinct TA batch numbers
            cursor.execute("SELECT DISTINCT ta_batch_number FROM ta WHERE ta_batch_number IS NOT NULL ORDER BY ta_batch_number")
            ta_batches = [row['ta_batch_number'] for row in cursor.fetchall()]
            
            # Get distinct trades
            cursor.execute("SELECT DISTINCT trade FROM ta WHERE trade IS NOT NULL ORDER BY trade")
            trades = [row['trade'] for row in cursor.fetchall()]
            
            # Get distinct final_result values
            cursor.execute("SELECT DISTINCT final_result FROM ta WHERE final_result IS NOT NULL ORDER BY final_result")
            final_results = [row['final_result'] for row in cursor.fetchall()]
            
            # Get distinct training names
            cursor.execute("SELECT DISTINCT training_name FROM ta WHERE training_name IS NOT NULL ORDER BY training_name")
            training_names = [row['training_name'] for row in cursor.fetchall()]
            
        return jsonify({
            'genders': genders,
            'academic_years': academic_years,
            'ta_batches': ta_batches,
            'trades': trades,
            'final_results': final_results,
            'training_names': training_names
        })
    except Exception as e:
        print(f"Error fetching TA filter options: {e}")
        return jsonify({
            'genders': [],
            'academic_years': [],
            'ta_batches': [],
            'trades': [],
            'final_results': [],
            'training_names': []
        }), 500
    finally:
        conn.close()

# TA Data API
@user_tech_bp.route('/api/ta/data', methods=['GET'])
def get_ta_data():
    try:
        # Get filter parameters from request
        gender = request.args.get('gender', 'all')
        academic_year = request.args.get('academic_year', 'all')
        ta_batch = request.args.get('ta_batch', 'all')
        trade = request.args.get('trade', 'all')
        final_result = request.args.get('final_result', 'all')
        training_name = request.args.get('training_name', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        # Pagination parameters
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            # Handle academic year filter
            if academic_year != 'all':
                # Extract the base year from the formatted academic year (e.g., "2023/24" -> 2023)
                base_year = academic_year.split('/')[0]
                where_conditions.append("joining_year = %s")
                params.append(base_year)
            
            if ta_batch != 'all':
                where_conditions.append("ta_batch_number = %s")
                params.append(ta_batch)
            
            if trade != 'all':
                where_conditions.append("trade = %s")
                params.append(trade)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training_name != 'all':
                where_conditions.append("training_name = %s")
                params.append(training_name)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_separation <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM ta WHERE {where_clause} ORDER BY sr_no DESC LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Get total count for pagination
            count_sql = f"SELECT COUNT(*) as total FROM ta WHERE {where_clause}"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()['total']
            
            # Calculate statistics with additional fields
            stats_sql = f"""
                SELECT 
                    COUNT(DISTINCT ta_batch_number) as batch_count,
                    COUNT(*) as coverage_count,
                    COUNT(CASE WHEN gender = 'Male' THEN 1 END) as male_count,
                    COUNT(CASE WHEN gender = 'Female' THEN 1 END) as female_count,
                    COUNT(CASE WHEN date_of_separation IS NULL OR date_of_separation > CURDATE() THEN 1 END) as live_count,
                    COUNT(CASE WHEN date_of_separation IS NOT NULL AND date_of_separation <= CURDATE() THEN 1 END) as separated_count
                FROM ta 
                WHERE {where_clause}
            """
            cursor.execute(stats_sql, params)
            stats = cursor.fetchone()
            
        # Convert records to list of dictionaries for JSON serialization
        records_list = []
        for record in records:
            # Format academic year for display
            academic_year_display = ""
            if record['joining_year']:
                academic_year_display = f"{int(record['joining_year'])}/{str(int(record['joining_year'])+1)[2:]}"
            
            records_list.append({
                'sr_no': record['sr_no'],
                'ticket_no': record['ticket_no'],
                'name': record['name'],
                'gender': record['gender'],
                'academic_year': academic_year_display,
                'date_of_joining': record['date_of_joining'].strftime('%Y-%m-%d') if record['date_of_joining'] else '',
                'ta_batch_number': record['ta_batch_number'],
                'date_of_separation': record['date_of_separation'].strftime('%Y-%m-%d') if record['date_of_separation'] else '',
                'trade': record['trade'],
                'final_result': record['final_result'],
                'training_name': record['training_name']
            })
        
        return jsonify({
            'records': records_list,
            'stats': {
                'batch_count': stats['batch_count'] if stats else 0,
                'coverage_count': stats['coverage_count'] if stats else 0,
                'male_count': stats['male_count'] if stats else 0,
                'female_count': stats['female_count'] if stats else 0,
                'live_count': stats['live_count'] if stats else 0,
                'separated_count': stats['separated_count'] if stats else 0
            },
            'total_records': total_count
        })
        
    except Exception as e:
        print(f"Error fetching TA data: {e}")
        return jsonify({
            'records': [],
            'stats': {
                'batch_count': 0,
                'coverage_count': 0,
                'male_count': 0,
                'female_count': 0,
                'live_count': 0,
                'separated_count': 0
            },
            'total_records': 0
        }), 500
    finally:
        conn.close()

# TA Download Excel API
@user_tech_bp.route('/api/ta/download', methods=['GET'])
def download_ta_data():
    try:
        # Get filter parameters from request
        gender = request.args.get('gender', 'all')
        academic_year = request.args.get('academic_year', 'all')
        ta_batch = request.args.get('ta_batch', 'all')
        trade = request.args.get('trade', 'all')
        final_result = request.args.get('final_result', 'all')
        training_name = request.args.get('training_name', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            # Handle academic year filter
            if academic_year != 'all':
                base_year = academic_year.split('/')[0]
                where_conditions.append("joining_year = %s")
                params.append(base_year)
            
            if ta_batch != 'all':
                where_conditions.append("ta_batch_number = %s")
                params.append(ta_batch)
            
            if trade != 'all':
                where_conditions.append("trade = %s")
                params.append(trade)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training_name != 'all':
                where_conditions.append("training_name = %s")
                params.append(training_name)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_separation <= %s")
                params.append(end_date)
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM ta WHERE {where_clause} ORDER BY sr_no DESC"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Convert records to DataFrame
            df = pd.DataFrame(records)
            
            # Rename columns to proper names
            column_mapping = {
                'sr_no': 'Sr No',
                'ticket_no': 'Ticket No',
                'name': 'Name',
                'gender': 'Gender',
                'joining_year': 'Joining Year',
                'date_of_joining': 'Date of Joining',
                'ta_batch_number': 'TA Batch Number',
                'date_of_separation': 'Date of Separation',
                'trade': 'Trade',
                'final_result': 'Final Result',
                'training_name': 'Training Name'
            }
            
            # Rename the columns
            df = df.rename(columns=column_mapping)
            
            # Create Academic Year column from joining_year and format it
            if 'Joining Year' in df.columns:
                df['Academic Year'] = df['Joining Year'].apply(
                    lambda x: f"{int(x)}/{str(int(x)+1)[2:]}" if pd.notnull(x) else ''
                )
                # Drop the original joining_year column
                df = df.drop(columns=['Joining Year'])
            
            # Format dates
            date_columns = ['Date of Joining', 'Date of Separation']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
            
            # Reorder columns to match the desired order
            desired_columns = [
                'Sr No', 'Ticket No', 'Name', 'Gender', 'Academic Year', 
                'Date of Joining', 'TA Batch Number', 'Date of Separation', 
                'Trade', 'Final Result', 'Training Name'
            ]
            
            # Filter to only include columns that exist in the dataframe
            columns_to_use = [col for col in desired_columns if col in df.columns]
            df = df[columns_to_use]
            
            # Create an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='TA Data')
                
                # Adjust column widths for better readability
                worksheet = writer.sheets['TA Data']
                for idx, col in enumerate(df.columns):
                    # Set column width based on content length
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2  # Add some padding
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)  # Cap width at 50
            
            output.seek(0)
            
            # Return the Excel file as a response
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='ta_data.xlsx'
            )
            
    except Exception as e:
        print(f"Error generating TA Excel file: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()
# Kaushalya Main Page
@user_tech_bp.route('/kaushalya', methods=['GET'])
def kaushalya_list():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT * FROM kaushalya ORDER BY sr_no DESC"
            cursor.execute(sql)
            kaushalya_data = cursor.fetchall()
    except Exception as e:
        print(f"Error fetching Kaushalya data: {e}")
        kaushalya_data = []
    finally:
        conn.close()

    return render_template("user/kaushalya.html", kaushalya_data=kaushalya_data)

# Kaushalya Filter Options API
@user_tech_bp.route('/api/kaushalya/filter-options', methods=['GET'])
def get_kaushalya_filter_options():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Get unique values for each filter
            cursor.execute("SELECT DISTINCT trade FROM kaushalya WHERE trade IS NOT NULL ORDER BY trade")
            trades = [row['trade'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT kaushalya_batch_no FROM kaushalya WHERE kaushalya_batch_no IS NOT NULL ORDER BY kaushalya_batch_no")
            batches = [row['kaushalya_batch_no'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT gender FROM kaushalya WHERE gender IS NOT NULL ORDER BY gender")
            genders = [row['gender'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT dei_batch FROM kaushalya WHERE dei_batch IS NOT NULL ORDER BY dei_batch")
            dei_batches = [row['dei_batch'] for row in cursor.fetchall()]
            
            # Get distinct joining years
            cursor.execute("SELECT DISTINCT joining_year FROM kaushalya WHERE joining_year IS NOT NULL ORDER BY joining_year")
            joining_years = [str(row['joining_year']) for row in cursor.fetchall()]
            
            # Get distinct final results
            cursor.execute("SELECT DISTINCT final_result FROM kaushalya WHERE final_result IS NOT NULL ORDER BY final_result")
            final_results = [row['final_result'] for row in cursor.fetchall()]
            
            # Get distinct training names
            cursor.execute("SELECT DISTINCT training_name FROM kaushalya WHERE training_name IS NOT NULL ORDER BY training_name")
            trainings = [row['training_name'] for row in cursor.fetchall()]
            
            # Get distinct placement drives
            cursor.execute("SELECT DISTINCT placement_drive FROM kaushalya WHERE placement_drive IS NOT NULL ORDER BY placement_drive")
            placement_drives = [row['placement_drive'] for row in cursor.fetchall()]
            
        return jsonify({
            'trades': trades,
            'batches': batches,
            'genders': genders,
            'dei_batches': dei_batches,
            'joining_years': joining_years,
            'final_results': final_results,
            'trainings': trainings,
            'placement_drives': placement_drives
        })
    except Exception as e:
        print(f"Error fetching Kaushalya filter options: {e}")
        return jsonify({
            'trades': [],
            'batches': [],
            'genders': [],
            'dei_batches': [],
            'joining_years': [],
            'final_results': [],
            'trainings': [],
            'placement_drives': []
        }), 500
    finally:
        conn.close()

# Kaushalya Data API
@user_tech_bp.route('/api/kaushalya/data', methods=['GET'])
def get_kaushalya_data():
    try:
        # Get filter parameters from request
        trade = request.args.get('trade', 'all')
        batch = request.args.get('batch', 'all')
        gender = request.args.get('gender', 'all')
        dei_batch = request.args.get('dei_batch', 'all')
        joining_year = request.args.get('joining_year', 'all')
        final_result = request.args.get('final_result', 'all')
        training = request.args.get('training', 'all')
        placement_drive = request.args.get('placement_drive', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        tab = request.args.get('tab', 'overall')
        
        # Pagination parameters
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if trade != 'all':
                where_conditions.append("trade = %s")
                params.append(trade)
            
            if batch != 'all':
                where_conditions.append("kaushalya_batch_no = %s")
                params.append(batch)
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if dei_batch != 'all':
                where_conditions.append("dei_batch = %s")
                params.append(dei_batch)
            
            if joining_year != 'all':
                where_conditions.append("joining_year = %s")
                params.append(joining_year)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training != 'all':
                where_conditions.append("training_name = %s")
                params.append(training)
            
            if placement_drive != 'all':
                where_conditions.append("placement_drive = %s")
                params.append(placement_drive)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_joining <= %s")
                params.append(end_date)
            
            # Add tab-specific conditions
            if tab == 'live':
                # Students whose final result is not yet declared
                where_conditions.append("(final_result IS NULL OR final_result = '' OR final_result IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester'))")
            elif tab == 'sem1':
                # Students in semester 1: 
                # sem_1_pass_fail is not 'Pass' or 'Fail' (it's pending, incomplete, or empty)
                where_conditions.append("(sem_1_pass_fail IS NULL OR sem_1_pass_fail = '' OR sem_1_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester'))")
            elif tab == 'sem2':
                # Students in semester 2:
                # sem_1_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_2_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_1_pass_fail IN ('Pass', 'Fail') AND (sem_2_pass_fail IS NULL OR sem_2_pass_fail = '' OR sem_2_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'sem3':
                # Students in semester 3:
                # sem_2_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_3_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_2_pass_fail IN ('Pass', 'Fail') AND (sem_3_pass_fail IS NULL OR sem_3_pass_fail = '' OR sem_3_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'sem4':
                # Students in semester 4:
                # sem_3_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_4_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_3_pass_fail IN ('Pass', 'Fail') AND (sem_4_pass_fail IS NULL OR sem_4_pass_fail = '' OR sem_4_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'sem5':
                # Students in semester 5:
                # sem_4_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_5_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_4_pass_fail IN ('Pass', 'Fail') AND (sem_5_pass_fail IS NULL OR sem_5_pass_fail = '' OR sem_5_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'sem6':
                # Students in semester 6:
                # sem_5_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_6_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_5_pass_fail IN ('Pass', 'Fail') AND (sem_6_pass_fail IS NULL OR sem_6_pass_fail = '' OR sem_6_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM kaushalya WHERE {where_clause} ORDER BY sr_no DESC LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Get total count for pagination
            count_sql = f"SELECT COUNT(*) as total FROM kaushalya WHERE {where_clause}"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()['total']
            
            # Calculate statistics
            stats_sql = f"""
                SELECT 
                    COUNT(DISTINCT kaushalya_batch_no) as batch_count,
                    COUNT(*) as coverage_count,
                    COUNT(CASE WHEN gender = 'Male' THEN 1 END) as male_count,
                    COUNT(CASE WHEN gender = 'Female' THEN 1 END) as female_count,
                    COUNT(CASE WHEN final_result = 'Pass' THEN 1 END) as pass_count,
                    COUNT(CASE WHEN final_result = 'Fail' THEN 1 END) as fail_count,
                    COUNT(CASE WHEN placement_drive = 'Placed' THEN 1 END) as placed_count
                FROM kaushalya 
                WHERE {where_clause}
            """
            cursor.execute(stats_sql, params)
            stats = cursor.fetchone()
            
        # Convert records to list of dictionaries for JSON serialization
        records_list = []
        for record in records:
            records_list.append({
                'sr_no': record['sr_no'],
                'ticket_no': record['ticket_no'],
                'name': record['name'],
                'gender': record['gender'],
                'joining_year': record['joining_year'],
                'date_of_joining': record['date_of_joining'].strftime('%Y-%m-%d') if record['date_of_joining'] else '',
                'kaushalya_batch_no': record['kaushalya_batch_no'],
                'trade': record['trade'],
                'dei_batch': record['dei_batch'],
                'sem_1_pass_fail': record['sem_1_pass_fail'],
                'sem_2_pass_fail': record['sem_2_pass_fail'],
                'sem_3_pass_fail': record['sem_3_pass_fail'],
                'sem_4_pass_fail': record['sem_4_pass_fail'],
                'sem_5_pass_fail': record['sem_5_pass_fail'],
                'sem_6_pass_fail': record['sem_6_pass_fail'],
                'final_result': record['final_result'],
                'placement_drive': record['placement_drive'],
                'training_name': record['training_name'],
                'remark': record['remark']
            })
        
        return jsonify({
            'records': records_list,
            'stats': {
                'batch_count': stats['batch_count'] if stats else 0,
                'coverage_count': stats['coverage_count'] if stats else 0,
                'male_count': stats['male_count'] if stats else 0,
                'female_count': stats['female_count'] if stats else 0,
                'pass_count': stats['pass_count'] if stats else 0,
                'fail_count': stats['fail_count'] if stats else 0,
                'placed_count': stats['placed_count'] if stats else 0
            },
            'total_records': total_count
        })
        
    except Exception as e:
        print(f"Error fetching Kaushalya data: {e}")
        return jsonify({
            'records': [],
            'stats': {
                'batch_count': 0,
                'coverage_count': 0,
                'male_count': 0,
                'female_count': 0,
                'pass_count': 0,
                'fail_count': 0,
                'placed_count': 0
            },
            'total_records': 0
        }), 500
    finally:
        conn.close()

# Kaushalya Download Excel API
@user_tech_bp.route('/api/kaushalya/download', methods=['GET'])
def download_kaushalya_data():
    try:
        # Get filter parameters from request
        trade = request.args.get('trade', 'all')
        batch = request.args.get('batch', 'all')
        gender = request.args.get('gender', 'all')
        dei_batch = request.args.get('dei_batch', 'all')
        joining_year = request.args.get('joining_year', 'all')
        final_result = request.args.get('final_result', 'all')
        training = request.args.get('training', 'all')
        placement_drive = request.args.get('placement_drive', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        tab = request.args.get('tab', 'overall')
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if trade != 'all':
                where_conditions.append("trade = %s")
                params.append(trade)
            
            if batch != 'all':
                where_conditions.append("kaushalya_batch_no = %s")
                params.append(batch)
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if dei_batch != 'all':
                where_conditions.append("dei_batch = %s")
                params.append(dei_batch)
            
            if joining_year != 'all':
                where_conditions.append("joining_year = %s")
                params.append(joining_year)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training != 'all':
                where_conditions.append("training_name = %s")
                params.append(training)
            
            if placement_drive != 'all':
                where_conditions.append("placement_drive = %s")
                params.append(placement_drive)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_joining <= %s")
                params.append(end_date)
            
            # Add tab-specific conditions
            if tab == 'live':
                # Students whose final result is not yet declared
                where_conditions.append("(final_result IS NULL OR final_result = '' OR final_result IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester'))")
            elif tab == 'sem1':
                # Students in semester 1: 
                # sem_1_pass_fail is not 'Pass' or 'Fail' (it's pending, incomplete, or empty)
                where_conditions.append("(sem_1_pass_fail IS NULL OR sem_1_pass_fail = '' OR sem_1_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester'))")
            elif tab == 'sem2':
                # Students in semester 2:
                # sem_1_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_2_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_1_pass_fail IN ('Pass', 'Fail') AND (sem_2_pass_fail IS NULL OR sem_2_pass_fail = '' OR sem_2_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'sem3':
                # Students in semester 3:
                # sem_2_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_3_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_2_pass_fail IN ('Pass', 'Fail') AND (sem_3_pass_fail IS NULL OR sem_3_pass_fail = '' OR sem_3_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'sem4':
                # Students in semester 4:
                # sem_3_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_4_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_3_pass_fail IN ('Pass', 'Fail') AND (sem_4_pass_fail IS NULL OR sem_4_pass_fail = '' OR sem_4_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'sem5':
                # Students in semester 5:
                # sem_4_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_5_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_4_pass_fail IN ('Pass', 'Fail') AND (sem_5_pass_fail IS NULL OR sem_5_pass_fail = '' OR sem_5_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'sem6':
                # Students in semester 6:
                # sem_5_pass_fail is 'Pass' or 'Fail' (completed) AND
                # sem_6_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(sem_5_pass_fail IN ('Pass', 'Fail') AND (sem_6_pass_fail IS NULL OR sem_6_pass_fail = '' OR sem_6_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM kaushalya WHERE {where_clause} ORDER BY sr_no DESC"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Convert records to DataFrame
            df = pd.DataFrame(records)
            
            # Rename columns to proper names
            column_mapping = {
                'sr_no': 'Sr No',
                'ticket_no': 'Ticket No',
                'name': 'Name',
                'gender': 'Gender',
                'joining_year': 'Joining Year',
                'date_of_joining': 'Date of Joining',
                'kaushalya_batch_no': 'Kaushalya Batch No',
                'trade': 'Trade',
                'dei_batch': 'DEI Batch',
                'sem_1_pass_fail': 'Sem 1 Pass/Fail',
                'sem_2_pass_fail': 'Sem 2 Pass/Fail',
                'sem_3_pass_fail': 'Sem 3 Pass/Fail',
                'sem_4_pass_fail': 'Sem 4 Pass/Fail',
                'sem_5_pass_fail': 'Sem 5 Pass/Fail',
                'sem_6_pass_fail': 'Sem 6 Pass/Fail',
                'final_result': 'Final Result',
                'placement_drive': 'Placement Drive',
                'training_name': 'Training Name',
                'remark': 'Remark'
            }
            
            # Rename the columns
            df = df.rename(columns=column_mapping)
            
            # Format dates
            date_columns = ['Date of Joining']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
            
            # Reorder columns to match the desired order
            desired_columns = [
                'Sr No', 'Ticket No', 'Name', 'Gender', 'Joining Year', 
                'Date of Joining', 'Kaushalya Batch No', 'Trade', 'DEI Batch',
                'Sem 1 Pass/Fail', 'Sem 2 Pass/Fail', 'Sem 3 Pass/Fail', 
                'Sem 4 Pass/Fail', 'Sem 5 Pass/Fail', 'Sem 6 Pass/Fail',
                'Final Result', 'Placement Drive', 'Training Name', 'Remark'
            ]
            
            # Filter to only include columns that exist in the dataframe
            columns_to_use = [col for col in desired_columns if col in df.columns]
            df = df[columns_to_use]
            
            # Create an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Kaushalya Data')
                
                # Adjust column widths for better readability
                worksheet = writer.sheets['Kaushalya Data']
                for idx, col in enumerate(df.columns):
                    # Set column width based on content length
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2  # Add some padding
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)  # Cap width at 50
            
            output.seek(0)
            
            # Return the Excel file as a response
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='kaushalya_data.xlsx'
            )
            
    except Exception as e:
        print(f"Error generating Kaushalya Excel file: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()
# Live Trainer Main Page
@user_tech_bp.route('/live_trainer', methods=['GET'])
def live_trainer_list():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT * FROM live_trainer ORDER BY sr_no DESC"
            cursor.execute(sql)
            live_trainer_data = cursor.fetchall()
    except Exception as e:
        print(f"Error fetching live trainer data: {e}")
        live_trainer_data = []
    finally:
        conn.close()

    return render_template("user/live_trainer.html", live_trainer_data=live_trainer_data)

# Live Trainer Filter Options API
@user_tech_bp.route('/api/live_trainer/filter-options', methods=['GET'])
def get_live_trainer_filter_options():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Get unique values for each filter
            cursor.execute("SELECT DISTINCT area FROM live_trainer WHERE area IS NOT NULL ORDER BY area")
            areas = [row['area'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT dept FROM live_trainer WHERE dept IS NOT NULL ORDER BY dept")
            departments = [row['dept'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT factory FROM live_trainer WHERE factory IS NOT NULL ORDER BY factory")
            factories = [row['factory'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT expertise_area FROM live_trainer WHERE expertise_area IS NOT NULL ORDER BY expertise_area")
            expertise_areas = [row['expertise_area'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT expertise_category FROM live_trainer WHERE expertise_category IS NOT NULL ORDER BY expertise_category")
            expertise_categories = [row['expertise_category'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT faculty_name FROM live_trainer WHERE faculty_name IS NOT NULL ORDER BY faculty_name")
            faculties = [row['faculty_name'] for row in cursor.fetchall()]
            
        return jsonify({
            'areas': areas,
            'departments': departments,
            'factories': factories,
            'expertise_areas': expertise_areas,
            'expertise_categories': expertise_categories,
            'faculties': faculties
        })
    except Exception as e:
        print(f"Error fetching live trainer filter options: {e}")
        return jsonify({
            'areas': [],
            'departments': [],
            'factories': [],
            'expertise_areas': [],
            'expertise_categories': [],
            'faculties': []
        }), 500
    finally:
        conn.close()

# Live Trainer Data API
@user_tech_bp.route('/api/live_trainer/data', methods=['GET'])
def get_live_trainer_data():
    try:
        # Get filter parameters from request
        area = request.args.get('area', 'all')
        department = request.args.get('department', 'all')
        factory = request.args.get('factory', 'all')
        expertise_area = request.args.get('expertise_area', 'all')
        expertise_category = request.args.get('expertise_category', 'all')
        faculty = request.args.get('faculty', 'all')
        ticket = request.args.get('ticket', '')
        
        # Pagination parameters
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if area != 'all':
                where_conditions.append("area = %s")
                params.append(area)
            
            if department != 'all':
                where_conditions.append("dept = %s")
                params.append(department)
            
            if factory != 'all':
                where_conditions.append("factory = %s")
                params.append(factory)
            
            if expertise_area != 'all':
                where_conditions.append("expertise_area = %s")
                params.append(expertise_area)
            
            if expertise_category != 'all':
                where_conditions.append("expertise_category = %s")
                params.append(expertise_category)
            
            if faculty != 'all':
                where_conditions.append("faculty_name = %s")
                params.append(faculty)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM live_trainer WHERE {where_clause} ORDER BY sr_no DESC LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Get total count for pagination
            count_sql = f"SELECT COUNT(*) as total FROM live_trainer WHERE {where_clause}"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()['total']
            
            # Calculate statistics - including area-based counts
            stats_sql = f"""
                SELECT 
                    COUNT(*) as total_count,
                    COUNT(CASE WHEN area = 'SDC Internal' THEN 1 END) as sdc_internal_count,
                    COUNT(CASE WHEN area = 'SDC Third Party' THEN 1 END) as sdc_third_party_count,
                    COUNT(CASE WHEN area = 'Shop Support' THEN 1 END) as shop_support_count
                FROM live_trainer 
                WHERE {where_clause}
            """
            cursor.execute(stats_sql, params)
            stats = cursor.fetchone()
            
        # Convert records to list of dictionaries for JSON serialization
        records_list = []
        for record in records:
            records_list.append({
                'sr_no': record['sr_no'],
                'faculty_name': record['faculty_name'],
                'ticket_no': record['ticket_no'],
                'mail_id': record['mail_id'],
                'mobile_number': record['mobile_number'],
                'area': record['area'],
                'dept': record['dept'],
                'factory': record['factory'],
                'reporting_manager_name': record['reporting_manager_name'],
                'reporting_manager_mail_id': record['reporting_manager_mail_id'],
                'expertise_area': record['expertise_area'],
                'expertise_category': record['expertise_category'],
                'hr_coordinator_name': record['hr_coordinator_name'],
                'remark': record['remark']
            })
        
        return jsonify({
            'records': records_list,
            'stats': {
                'total_count': stats['total_count'] if stats else 0,
                'sdc_internal_count': stats['sdc_internal_count'] if stats else 0,
                'sdc_third_party_count': stats['sdc_third_party_count'] if stats else 0,
                'shop_support_count': stats['shop_support_count'] if stats else 0
            },
            'total_records': total_count
        })
        
    except Exception as e:
        print(f"Error fetching live trainer data: {e}")
        return jsonify({
            'records': [],
            'stats': {
                'total_count': 0,
                'sdc_internal_count': 0,
                'sdc_third_party_count': 0,
                'shop_support_count': 0
            },
            'total_records': 0
        }), 500
    finally:
        conn.close()

# Live Trainer Download Excel API
@user_tech_bp.route('/api/live_trainer/download', methods=['GET'])
def download_live_trainer_data():
    try:
        # Get filter parameters from request
        area = request.args.get('area', 'all')
        department = request.args.get('department', 'all')
        factory = request.args.get('factory', 'all')
        expertise_area = request.args.get('expertise_area', 'all')
        expertise_category = request.args.get('expertise_category', 'all')
        faculty = request.args.get('faculty', 'all')
        ticket = request.args.get('ticket', '')
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if area != 'all':
                where_conditions.append("area = %s")
                params.append(area)
            
            if department != 'all':
                where_conditions.append("dept = %s")
                params.append(department)
            
            if factory != 'all':
                where_conditions.append("factory = %s")
                params.append(factory)
            
            if expertise_area != 'all':
                where_conditions.append("expertise_area = %s")
                params.append(expertise_area)
            
            if expertise_category != 'all':
                where_conditions.append("expertise_category = %s")
                params.append(expertise_category)
            
            if faculty != 'all':
                where_conditions.append("faculty_name = %s")
                params.append(faculty)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM live_trainer WHERE {where_clause} ORDER BY sr_no DESC"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Convert records to DataFrame
            df = pd.DataFrame(records)
            
            # Rename columns to proper names
            column_mapping = {
                'sr_no': 'Sr No',
                'faculty_name': 'Faculty Name',
                'ticket_no': 'Ticket No',
                'mail_id': 'Mail ID',
                'mobile_number': 'Mobile Number',
                'area': 'Area',
                'dept': 'Department',
                'factory': 'Factory',
                'reporting_manager_name': 'Reporting Manager Name',
                'reporting_manager_mail_id': 'Reporting Manager Mail ID',
                'expertise_area': 'Expertise Area',
                'expertise_category': 'Expertise Category',
                'hr_coordinator_name': 'HR Coordinator Name',
                'remark': 'Remark'
            }
            
            # Rename the columns
            df = df.rename(columns=column_mapping)
            
            # Reorder columns to match the desired order
            desired_columns = [
                'Sr No', 'Faculty Name', 'Ticket No', 'Mail ID', 'Mobile Number',
                'Area', 'Department', 'Factory', 'Reporting Manager Name',
                'Reporting Manager Mail ID', 'Expertise Area', 'Expertise Category',
                'HR Coordinator Name', 'Remark'
            ]
            
            # Filter to only include columns that exist in the dataframe
            columns_to_use = [col for col in desired_columns if col in df.columns]
            df = df[columns_to_use]
            
            # Create an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Live Trainer Data')
                
                # Adjust column widths for better readability
                worksheet = writer.sheets['Live Trainer Data']
                for idx, col in enumerate(df.columns):
                    # Set column width based on content length
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2  # Add some padding
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)  # Cap width at 50
            
            output.seek(0)
            
            # Return the Excel file as a response
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='live_trainer_data.xlsx'
            )
            
    except Exception as e:
        print(f"Error generating Live Trainer Excel file: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()        
# Lakshya Main Page
@user_tech_bp.route('/lakshya', methods=['GET'])
def lakshya_list():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT * FROM lakshya ORDER BY sr_no DESC"
            cursor.execute(sql)
            lakshya_data = cursor.fetchall()
    except Exception as e:
        print(f"Error fetching Lakshya data: {e}")
        lakshya_data = []
    finally:
        conn.close()

    return render_template("user/lakshya.html", lakshya_data=lakshya_data)

# Lakshya Filter Options API
@user_tech_bp.route('/api/lakshya/filter-options', methods=['GET'])
def get_lakshya_filter_options():
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Get unique values for each filter
            cursor.execute("SELECT DISTINCT diploma_name FROM lakshya WHERE diploma_name IS NOT NULL ORDER BY diploma_name")
            diplomas = [row['diploma_name'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT lakshya_batch_no FROM lakshya WHERE lakshya_batch_no IS NOT NULL ORDER BY lakshya_batch_no")
            batches = [row['lakshya_batch_no'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT gender FROM lakshya WHERE gender IS NOT NULL ORDER BY gender")
            genders = [row['gender'] for row in cursor.fetchall()]
            
            cursor.execute("SELECT DISTINCT diploma_trainee_inplant_shop FROM lakshya WHERE diploma_trainee_inplant_shop IS NOT NULL ORDER BY diploma_trainee_inplant_shop")
            inplant_shops = [row['diploma_trainee_inplant_shop'] for row in cursor.fetchall()]
            
            # Get distinct course joining years
            cursor.execute("SELECT DISTINCT course_joining_year FROM lakshya WHERE course_joining_year IS NOT NULL ORDER BY course_joining_year")
            joining_years = [str(row['course_joining_year']) for row in cursor.fetchall()]
            
            # Get distinct final results
            cursor.execute("SELECT DISTINCT final_result FROM lakshya WHERE final_result IS NOT NULL ORDER BY final_result")
            final_results = [row['final_result'] for row in cursor.fetchall()]
            
            # Get distinct training names
            cursor.execute("SELECT DISTINCT training_name FROM lakshya WHERE training_name IS NOT NULL ORDER BY training_name")
            trainings = [row['training_name'] for row in cursor.fetchall()]
            
        return jsonify({
            'diplomas': diplomas,
            'batches': batches,
            'genders': genders,
            'inplant_shops': inplant_shops,
            'joining_years': joining_years,
            'final_results': final_results,
            'trainings': trainings
        })
    except Exception as e:
        print(f"Error fetching Lakshya filter options: {e}")
        return jsonify({
            'diplomas': [],
            'batches': [],
            'genders': [],
            'inplant_shops': [],
            'joining_years': [],
            'final_results': [],
            'trainings': []
        }), 500
    finally:
        conn.close()

# Lakshya Data API
@user_tech_bp.route('/api/lakshya/data', methods=['GET'])
def get_lakshya_data():
    try:
        # Get filter parameters from request
        diploma = request.args.get('diploma', 'all')
        batch = request.args.get('batch', 'all')
        gender = request.args.get('gender', 'all')
        inplant_shop = request.args.get('inplant_shop', 'all')
        joining_year = request.args.get('joining_year', 'all')
        final_result = request.args.get('final_result', 'all')
        training = request.args.get('training', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        tab = request.args.get('tab', 'overall')
        
        # Pagination parameters
        limit = int(request.args.get('limit', 50))
        offset = int(request.args.get('offset', 0))
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if diploma != 'all':
                where_conditions.append("diploma_name = %s")
                params.append(diploma)
            
            if batch != 'all':
                where_conditions.append("lakshya_batch_no = %s")
                params.append(batch)
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if inplant_shop != 'all':
                where_conditions.append("diploma_trainee_inplant_shop = %s")
                params.append(inplant_shop)
            
            if joining_year != 'all':
                where_conditions.append("course_joining_year = %s")
                params.append(joining_year)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training != 'all':
                where_conditions.append("training_name = %s")
                params.append(training)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_joining <= %s")
                params.append(end_date)
            
            # Add tab-specific conditions
            if tab == 'live':
                # Students whose final result is not yet declared
                where_conditions.append("(final_result IS NULL OR final_result = '' OR final_result IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester'))")
            elif tab == 'sem1':
                # Students in semester 1: 
                # semester_1_pass_fail is not 'Pass' or 'Fail' (it's pending, incomplete, or empty)
                where_conditions.append("(semester_1_pass_fail IS NULL OR semester_1_pass_fail = '' OR semester_1_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester'))")
            elif tab == 'sem2':
                # Students in semester 2:
                # semester_1_pass_fail is 'Pass' or 'Fail' (completed) AND
                # semester_2_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(semester_1_pass_fail IN ('Pass', 'Fail') AND (semester_2_pass_fail IS NULL OR semester_2_pass_fail = '' OR semester_2_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'year2':
                # Students in second year:
                # semester_2_pass_fail is 'Pass' or 'Fail' (completed) AND
                # second_year_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(semester_2_pass_fail IN ('Pass', 'Fail') AND (second_year_pass_fail IS NULL OR second_year_pass_fail = '' OR second_year_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'year3':
                # Students in third year:
                # second_year_pass_fail is 'Pass' or 'Fail' (completed) AND
                # third_year_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(second_year_pass_fail IN ('Pass', 'Fail') AND (third_year_pass_fail IS NULL OR third_year_pass_fail = '' OR third_year_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'year4':
                # Students in fourth year:
                # third_year_pass_fail is 'Pass' or 'Fail' (completed) AND
                # fourth_year_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(third_year_pass_fail IN ('Pass', 'Fail') AND (fourth_year_pass_fail IS NULL OR fourth_year_pass_fail = '' OR fourth_year_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM lakshya WHERE {where_clause} ORDER BY sr_no DESC LIMIT {limit} OFFSET {offset}"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Get total count for pagination
            count_sql = f"SELECT COUNT(*) as total FROM lakshya WHERE {where_clause}"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()['total']
            
            # Calculate statistics
            stats_sql = f"""
                SELECT 
                    COUNT(DISTINCT lakshya_batch_no) as batch_count,
                    COUNT(*) as coverage_count,
                    COUNT(CASE WHEN gender = 'Male' THEN 1 END) as male_count,
                    COUNT(CASE WHEN gender = 'Female' THEN 1 END) as female_count,
                    COUNT(CASE WHEN final_result = 'Pass' THEN 1 END) as pass_count,
                    COUNT(CASE WHEN final_result = 'Fail' THEN 1 END) as fail_count
                FROM lakshya 
                WHERE {where_clause}
            """
            cursor.execute(stats_sql, params)
            stats = cursor.fetchone()
            
        # Convert records to list of dictionaries for JSON serialization
        records_list = []
        for record in records:
            records_list.append({
                'sr_no': record['sr_no'],
                'ticket_no': record['ticket_no'],
                'name': record['name'],
                'gender': record['gender'],
                'course_joining_year': record['course_joining_year'],
                'date_of_joining': record['date_of_joining'].strftime('%Y-%m-%d') if record['date_of_joining'] else '',
                'date_of_separation': record['date_of_separation'].strftime('%Y-%m-%d') if record['date_of_separation'] else '',
                'lakshya_batch_no': record['lakshya_batch_no'],
                'diploma_name': record['diploma_name'],
                'diploma_trainee_inplant_shop': record['diploma_trainee_inplant_shop'],
                'semester_1_pass_fail': record['semester_1_pass_fail'],
                'semester_2_pass_fail': record['semester_2_pass_fail'],
                'second_year_pass_fail': record['second_year_pass_fail'],
                'third_year_pass_fail': record['third_year_pass_fail'],
                'fourth_year_pass_fail': record['fourth_year_pass_fail'],
                'final_result': record['final_result'],
                'training_name': record['training_name'],
                'remark': record['remark']
            })
        
        return jsonify({
            'records': records_list,
            'stats': {
                'batch_count': stats['batch_count'] if stats else 0,
                'coverage_count': stats['coverage_count'] if stats else 0,
                'male_count': stats['male_count'] if stats else 0,
                'female_count': stats['female_count'] if stats else 0,
                'pass_count': stats['pass_count'] if stats else 0,
                'fail_count': stats['fail_count'] if stats else 0
            },
            'total_records': total_count
        })
        
    except Exception as e:
        print(f"Error fetching Lakshya data: {e}")
        return jsonify({
            'records': [],
            'stats': {
                'batch_count': 0,
                'coverage_count': 0,
                'male_count': 0,
                'female_count': 0,
                'pass_count': 0,
                'fail_count': 0
            },
            'total_records': 0
        }), 500
    finally:
        conn.close()

# Lakshya Download Excel API
@user_tech_bp.route('/api/lakshya/download', methods=['GET'])
def download_lakshya_data():
    try:
        # Get filter parameters from request
        diploma = request.args.get('diploma', 'all')
        batch = request.args.get('batch', 'all')
        gender = request.args.get('gender', 'all')
        inplant_shop = request.args.get('inplant_shop', 'all')
        joining_year = request.args.get('joining_year', 'all')
        final_result = request.args.get('final_result', 'all')
        training = request.args.get('training', 'all')
        ticket = request.args.get('ticket', '')
        start_date = request.args.get('startDate', '')
        end_date = request.args.get('endDate', '')
        tab = request.args.get('tab', 'overall')
        
        conn = get_db_connection()
        with conn.cursor() as cursor:
            # Build WHERE clause based on filters
            where_conditions = []
            params = []
            
            if diploma != 'all':
                where_conditions.append("diploma_name = %s")
                params.append(diploma)
            
            if batch != 'all':
                where_conditions.append("lakshya_batch_no = %s")
                params.append(batch)
            
            if gender != 'all':
                where_conditions.append("gender = %s")
                params.append(gender)
            
            if inplant_shop != 'all':
                where_conditions.append("diploma_trainee_inplant_shop = %s")
                params.append(inplant_shop)
            
            if joining_year != 'all':
                where_conditions.append("course_joining_year = %s")
                params.append(joining_year)
            
            if final_result != 'all':
                where_conditions.append("final_result = %s")
                params.append(final_result)
            
            if training != 'all':
                where_conditions.append("training_name = %s")
                params.append(training)
            
            if ticket:
                where_conditions.append("ticket_no LIKE %s")
                params.append(f"%{ticket}%")
            
            if start_date:
                where_conditions.append("date_of_joining >= %s")
                params.append(start_date)
            
            if end_date:
                where_conditions.append("date_of_joining <= %s")
                params.append(end_date)
            
            # Add tab-specific conditions
            if tab == 'live':
                # Students whose final result is not yet declared
                where_conditions.append("(final_result IS NULL OR final_result = '' OR final_result IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester'))")
            elif tab == 'sem1':
                # Students in semester 1: 
                # semester_1_pass_fail is not 'Pass' or 'Fail' (it's pending, incomplete, or empty)
                where_conditions.append("(semester_1_pass_fail IS NULL OR semester_1_pass_fail = '' OR semester_1_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester'))")
            elif tab == 'sem2':
                # Students in semester 2:
                # semester_1_pass_fail is 'Pass' or 'Fail' (completed) AND
                # semester_2_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(semester_1_pass_fail IN ('Pass', 'Fail') AND (semester_2_pass_fail IS NULL OR semester_2_pass_fail = '' OR semester_2_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'year2':
                # Students in second year:
                # semester_2_pass_fail is 'Pass' or 'Fail' (completed) AND
                # second_year_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(semester_2_pass_fail IN ('Pass', 'Fail') AND (second_year_pass_fail IS NULL OR second_year_pass_fail = '' OR second_year_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'year3':
                # Students in third year:
                # second_year_pass_fail is 'Pass' or 'Fail' (completed) AND
                # third_year_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(second_year_pass_fail IN ('Pass', 'Fail') AND (third_year_pass_fail IS NULL OR third_year_pass_fail = '' OR third_year_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            elif tab == 'year4':
                # Students in fourth year:
                # third_year_pass_fail is 'Pass' or 'Fail' (completed) AND
                # fourth_year_pass_fail is not 'Pass' or 'Fail' (pending, incomplete, or empty)
                where_conditions.append("(third_year_pass_fail IN ('Pass', 'Fail') AND (fourth_year_pass_fail IS NULL OR fourth_year_pass_fail = '' OR fourth_year_pass_fail IN ('Pending', 'Incomplete', 'Exam pending due to Incomplete Semester')))")
            
            # Build the SQL query
            where_clause = " AND ".join(where_conditions) if where_conditions else "1=1"
            sql = f"SELECT * FROM lakshya WHERE {where_clause} ORDER BY sr_no DESC"
            
            cursor.execute(sql, params)
            records = cursor.fetchall()
            
            # Convert records to DataFrame
            df = pd.DataFrame(records)
            
            # Rename columns to proper names
            column_mapping = {
                'sr_no': 'Sr No',
                'ticket_no': 'Ticket No',
                'name': 'Name',
                'gender': 'Gender',
                'course_joining_year': 'Course Joining Year',
                'date_of_joining': 'Date of Joining',
                'date_of_separation': 'Date of Separation',
                'lakshya_batch_no': 'Lakshya Batch No',
                'diploma_name': 'Diploma Name',
                'diploma_trainee_inplant_shop': 'Diploma Trainee Inplant Shop',
                'semester_1_pass_fail': 'Semester 1 Pass/Fail',
                'semester_2_pass_fail': 'Semester 2 Pass/Fail',
                'second_year_pass_fail': 'Second Year Pass/Fail',
                'third_year_pass_fail': 'Third Year Pass/Fail',
                'fourth_year_pass_fail': 'Fourth Year Pass/Fail',
                'final_result': 'Final Result',
                'training_name': 'Training Name',
                'remark': 'Remark'
            }
            
            # Rename the columns
            df = df.rename(columns=column_mapping)
            
            # Format dates
            date_columns = ['Date of Joining', 'Date of Separation']
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
            
            # Reorder columns to match the desired order
            desired_columns = [
                'Sr No', 'Ticket No', 'Name', 'Gender', 'Course Joining Year', 
                'Date of Joining', 'Date of Separation', 'Lakshya Batch No', 
                'Diploma Name', 'Diploma Trainee Inplant Shop',
                'Semester 1 Pass/Fail', 'Semester 2 Pass/Fail', 
                'Second Year Pass/Fail', 'Third Year Pass/Fail', 'Fourth Year Pass/Fail',
                'Final Result', 'Training Name', 'Remark'
            ]
            
            # Filter to only include columns that exist in the dataframe
            columns_to_use = [col for col in desired_columns if col in df.columns]
            df = df[columns_to_use]
            
            # Create an Excel file in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Lakshya Data')
                
                # Adjust column widths for better readability
                worksheet = writer.sheets['Lakshya Data']
                for idx, col in enumerate(df.columns):
                    # Set column width based on content length
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(col)
                    ) + 2  # Add some padding
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)  # Cap width at 50
            
            output.seek(0)
            
            # Return the Excel file as a response
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='lakshya_data.xlsx'
            )
            
    except Exception as e:
        print(f"Error generating Lakshya Excel file: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()