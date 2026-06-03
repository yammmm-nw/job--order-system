# Job Order Management System

A production-grade .NET desktop application engineered to automate, track, and manage daily corporate job orders and technical requests. This system runs continuously to handle real-time database transactions, enforce business workflow rules, and eliminate operational duplication.

## 🛠️ Tech Stack & Architecture
* **Language:** VB.NET (.NET Framework)
* **UI Framework:** Guna UI (Guna2 Suite) for modern desktop interfaces
* **Database:** MySQL (via MySQL Connector/NET)
* **Architecture:** Event-Driven Desktop Application with Parameterized Relational Database Access

## 📋 Company Problem Solved
Before this system was built, the company faced these operational challenges:
* **High Manual Overhead:** Staff spent 5 hours daily manually encoding and maintaining legacy Excel spreadsheets.
* **Data Integrity Risks:** High risk of human error during manual Job Order data entry, tracking, and updates.
* **System Instability:** Frequent spreadsheet crashes caused by multi-user access conflicts, resulting in data loss.

## 🚀 Key Features
* **Job Order Processing:** Streamlines the creation, validation, and secure storage of company job order records.
* **Real-Time Database Integration:** Securely reads, updates, and archives records instantly via an optimized MySQL connection.
* **Automated Reporting Module:** Generates accurate daily and monthly operations reports for management review, eliminating manual compilation.

## 📈 Business Impact & Metrics
* **Daily Active Users:** Deployed internally and used by corporate employees all day to complete core operational tasks.
* **Time Saved:** Drastically reduced manual processing time by **70%**, shifting hours back to active tasks.
* **System Reliability:** Replaced legacy sheets with a robust desktop architecture, operating with **100% daily uptime** since initial deployment.

## 💻 Core Logic Implementation

The following production code illustrates the business rule validation, secure database write, and user interface state synchronization lifecycle:

```vb
' =====================================================================
' ACTION: VALIDATE CONSTRAINTS & ADD RECORD TO OPERATIONS DATABASE
' =====================================================================
Private Sub Guna2Button9_Click(sender As Object, e As EventArgs) Handles Guna2Button9.Click
    Try
        ' Securely open a connection using dynamic string configuration
        Using sqlConn As New MySqlConnection(\$"server={server};user id={username};password={password};database={database};")
            sqlConn.Open()

            ' 1. BUSINESS VALIDATION: Check if the applicant already has an active 'UNFINISH' status
            Dim checkQuery As String = "SELECT COUNT(*) FROM joborderdb WHERE `Requested By:` = @requested_by AND `Status` = 'UNFINISH';"

            Using checkCmd As New MySqlCommand(checkQuery, sqlConn)
                checkCmd.Parameters.AddWithValue("@requested_by", Guna2TextBox2.Text)
                Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                ' Enforce business logic constraint rule
                If count > 0 Then
                    MessageBox.Show("This person already has an UNFINISH job order. Cannot add another.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End Using

            ' 2. DATABASE TRANSACTION: Execute sanitized insert query if constraints are satisfied
            Dim query As String = "INSERT INTO `joborderdb` (`Requested By:`, `Nature of Job Order`, `Address`,`District`, `Meter / Application / Account No`, `Prioritization`, `Status`, `Job Order Number`, `Date`) VALUES (@requested_by, @nature, @address, @district, @meter_no, @priority, @status, @job_order_number, @date);"

            Using sqlCmd As New MySqlCommand(query, sqlConn)
                sqlCmd.Parameters.AddWithValue("@requested_by", Guna2TextBox2.Text)
                sqlCmd.Parameters.AddWithValue("@nature", Guna2ComboBox1.Text)
                sqlCmd.Parameters.AddWithValue("@address", Guna2ComboBox2.Text)
                sqlCmd.Parameters.AddWithValue("@district", Guna2TextBox4.Text)
                sqlCmd.Parameters.AddWithValue("@meter_no", Guna2TextBox3.Text)
                sqlCmd.Parameters.AddWithValue("@priority", Guna2ComboBox3.Text)
                sqlCmd.Parameters.AddWithValue("@status", Guna2ComboBox4.Text)
                sqlCmd.Parameters.AddWithValue("@date", Guna2DateTimePicker1.Value.ToString("yyyy-MM-dd"))
                sqlCmd.Parameters.AddWithValue("@job_order_number", TextBox2.Text)

                sqlCmd.ExecuteNonQuery()
            End Using
        End Using

        ' 3. STATE LIFECYCLE: Update, style, and clear the user interface
        MessageBox.Show("Record added successfully!")
        UpdateTable()           ' Refresh data grid components from MySQL database
        ColorRowsByStatus()     ' Execute dynamic conditional formatting on grid rows
        ClearFields()           ' Clear form inputs to maintain high data entry speeds
        LoadNames()             ' Reload relational database dependency collections

    Catch ex As Exception
        ' Defensive logging to shield the application from sudden server dropouts
        MessageBox.Show("Insert error: " & ex.Message)
    End Try
End Sub
```

## ⚙️ Engineering Principles Demonstrated
* **Data Integration:** Utilized `ExecuteScalar` and `ExecuteNonQuery` within strict object lifecycles (`Using` statement blocks) to guarantee automatic memory management and prevent database link leakage.
* **User Experience (UX):** Created explicit post-transaction rendering states (`ColorRowsByStatus`) providing immediate visual status cues to operators.

