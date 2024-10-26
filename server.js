const express = require('express');
const mysql = require('mysql');
const multer = require('multer');
const cors = require('cors');
const fs = require('fs');
const app = express();
const xlsx = require('xlsx');
const moment = require('moment');  // Moment.js for handling dates

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// MySQL connection
const db = mysql.createConnection({
  host: 'localhost',
  user: 'root',
  password: '',
  database: 'crud_db'
});

db.connect(err => {
  if (err) {
    console.log('Error connecting to the database:', err);
  } else {
    console.log('Connected to MySQL database');
  }
});

// Multer setup for file uploads
const upload = multer({ dest: 'uploads/' });

// Function to convert Excel date serial to JS Date
const excelDateToJSDate = (serial) => {
  const startDate = new Date(1900, 0, 1);
  const resultDate = new Date(startDate.getTime() + ((serial - 1) * 24 * 60 * 60 * 1000));
  return resultDate;
};

// Route to handle XLS file upload
app.post('/upload-xls', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  try {
    const workbook = xlsx.readFile(req.file.path);
    const sheet_name = workbook.SheetNames[0];
    const sheet = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name]);

    console.log('Uploaded sheet data:', sheet);

    sheet.forEach(row => {
      const itemName = row.item_name;
      const employeeId = row.employee_id;
      const itemCode = row['item_code\r\n'] || row.item_code;
      const trimmedItemCode = itemCode ? itemCode.toString().trim() : null;

      const sql = 'INSERT INTO item_table (item_name, employee_id, item_code) VALUES (?, ?, ?)';
      db.query(sql, [itemName, employeeId, trimmedItemCode], (err, result) => {
        if (err) {
          console.error('Error inserting data:', err);
          return;
        }
        console.log('Data inserted successfully:', result);
      });
    });

    res.json({ message: 'File uploaded and data inserted successfully' });

  } catch (error) {
    console.error('Error processing XLS file:', error);
    res.status(500).json({ error: 'Error processing XLS file' });
  } finally {
    fs.unlink(req.file.path, (err) => {
      if (err) {
        console.error('Error deleting uploaded file:', err);
      } else {
        console.log('Uploaded file deleted');
      }
    });
  }
});

// Route to upload employee info
app.post('/upload-employee-info', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  try {
    const workbook = xlsx.readFile(req.file.path);
    const sheet_name = workbook.SheetNames[0];
    const sheet = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name]);

    console.log('Uploaded employee info data:', sheet);

    sheet.forEach(row => {
      const employeeNumber = row.employee_number;
      const employeeName = row.employee_name;
      const employeeSalary = row.employee_salary;
      const employeePosition = row.position;

      const sql = 'INSERT INTO employee_info (employee_number, employee_name, employee_salary, position) VALUES (?, ?, ?, ?)';
      db.query(sql, [employeeNumber, employeeName, employeeSalary, employeePosition], (err, result) => {
        if (err) {
          console.error('Error inserting data into employee_info:', err);
          return;
        }
        console.log('Data inserted into employee_info successfully:', result);
      });
    });

    res.json({ message: 'Employee info file uploaded and data inserted successfully' });

  } catch (error) {
    console.error('Error processing employee info XLS file:', error);
    res.status(500).json({ error: 'Error processing employee info XLS file' });
  } finally {
    fs.unlink(req.file.path, (err) => {
      if (err) {
        console.error('Error deleting uploaded file:', err);
      } else {
        console.log('Uploaded file deleted');
      }
    });
  }
});

// Route to search employee records by employee number
app.get('/employee-records/:employee_number', (req, res) => {
  const employeeNumber = req.params.employee_number;

  const sql = `
    SELECT DATE(employee_time_in) AS date, 
           MIN(TIME(employee_time_in)) AS earliest_time, 
           MAX(TIME(employee_time_in)) AS last_time, 
           TIMEDIFF(MAX(employee_time_in), MIN(employee_time_in)) AS time_difference
    FROM employee_time_logs
    WHERE employee_number = ?
    GROUP BY DATE(employee_time_in)
  `;

  db.query(sql, [employeeNumber], (err, results) => {
    if (err) {
      console.error('Database query error:', err);
      return res.status(500).json({ error: 'Database query error' });
    }
    res.json(results);
  });
});

// Department Management Routes

// Route to add a new department
app.post('/departments', (req, res) => {
  const { department_name, department_code } = req.body;

  if (!department_name || !department_code) {
    return res.status(400).json({ error: 'Department name and code are required' });
  }

  const sql = 'INSERT INTO department (department_name, department_code) VALUES (?, ?)';
  db.query(sql, [department_name, department_code], (err, result) => {
    if (err) {
      console.error('Error adding department:', err);
      return res.status(500).json({ error: 'Error adding department' });
    }
    res.json({ message: 'Department added successfully', departmentId: result.insertId });
  });
});

// Route to update an existing department
app.put('/departments/:id', (req, res) => {
  const { id } = req.params;
  const { department_name, department_code } = req.body;

  if (!department_name || !department_code) {
    return res.status(400).json({ error: 'Department name and code are required' });
  }

  const sql = 'UPDATE department SET department_name = ?, department_code = ? WHERE id = ?';
  db.query(sql, [department_name, department_code, id], (err, result) => {
    if (err) {
      console.error('Error updating department:', err);
      return res.status(500).json({ error: 'Error updating department' });
    }
    res.json({ message: 'Department updated successfully' });
  });
});

// Route to delete a department
app.delete('/departments/:id', (req, res) => {
  const { id } = req.params;

  const sql = 'DELETE FROM department WHERE id = ?';
  db.query(sql, [id], (err, result) => {
    if (err) {
      console.error('Error deleting department:', err);
      return res.status(500).json({ error: 'Error deleting department' });
    }
    res.json({ message: 'Department deleted successfully' });
  });
});

// Route to get all departments
app.get('/departments', (req, res) => {
  const sql = 'SELECT * FROM department';
  db.query(sql, (err, results) => {
    if (err) {
      console.error('Error fetching departments:', err);
      return res.status(500).json({ error: 'Error fetching departments' });
    }
    res.json(results);
  });
});

// Start the server
const port = 5000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
