using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using MetroFramework;
using MetroFramework.Forms;
using System.Configuration;
using HealthClaimApp.Utilities; // Assuming you have a namespace like this
using DocumentFormat.OpenXml.Spreadsheet;
using RJCodeAdvance.RJControls;
using System.Text;
using ClosedXML.Excel;
using Dapper;
using System.ComponentModel;
using DocumentFormat.OpenXml.Office.Word;
using MetroFramework.Components;
using System.Diagnostics;
using System.Globalization;
using static HealthClaimApp.Utilities.EmployeeAccount; // Adjust based on actual usage
using DataTable = System.Data.DataTable;
using ClosedXML.Excel.Exceptions;
using ADGV;
using System.Collections.Generic;

namespace HealthClaimApp
{
    public partial class frmClaimSubmission : MetroForm
    {
        private readonly string connectionString;
        private OleDbConnection connection;
        private DataTable claimDetails;
        private readonly EmployeeAccount employeeAccount;
        private string cultureCode;

        public frmClaimSubmission(string cultureCode)
        {
            InitializeComponent();
            connectionString = ConfigurationManager.ConnectionStrings["HealthClaimDB"].ConnectionString;
            this.cultureCode = cultureCode;
            SetCulture(cultureCode);
            InitializeEmployeeIDSelector();
            InitializeFamilyMemberIDSelector();
            hideEnableBtns();
            SetupDataGridView(); // Setup DataGridView columns
            LocalizeDataGridViewBenefitsViewHeaders(); // Localize DataGridView headers
            lblOn.Text = HealthClaimApp.Properties.Strings.EmployeeClaim;
            lblOff.Text = HealthClaimApp.Properties.Strings.FamilyMemberClaim;
            txtQuantity.Text = "1";
        }
        private void frmClaimSubmission_Load(object sender, EventArgs e)
        {
            PopulateComboBoxes();
            //SetupDataGridView();
            // PopulateBenefits("EL"); // Assuming English as default
            PopulateBenefits(cultureCode);
            LocalizeDataGridViewBenefitsViewHeaders(); // Ensure headers are localized
            cmbEmployee.Focus();

        }
        private void SetCulture(string cultureCode)
        {
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo(cultureCode);
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo(cultureCode);
            UpdateUIStrings();
            LocalizeDataGridViewBenefitsViewHeaders();
        }

        private void PopulateComboBoxes()
        {
            string employeeQuery = "SELECT EmployeeID, " + "[FirstName] & ' ' & [LastName] & ' (' & [EmployeeID] & ')' AS FullName " +
                          "FROM tblEmployees " + "WHERE IsActive = True " + "ORDER BY [FirstName] & ' ' & [LastName] ASC";
            string familyMemberQuery = "SELECT FamilyMemberID, [FirstName] & ' ' & [LastName] AS FullName " +
                                       "FROM tblFamilyMembers WHERE IsActive = True ORDER BY [FirstName] & ' ' & [LastName] ASC";

            PopulateComboBox(cmbEmployee, employeeQuery, "FullName", "EmployeeID");
            PopulateComboBox(cmbFamilyMember, familyMemberQuery, "FullName", "FamilyMemberID");
        }
        private void UpdateUIStrings()

        {
            // Update form strings based on selected language
            lblClaimSumTittle.Text = HealthClaimApp.Properties.Strings.LBLClaimSumTittle;
            cmbEmployee.Text = HealthClaimApp.Properties.Strings.cmbEMPLOYEE;
            cmbFamilyMember.Text = HealthClaimApp.Properties.Strings.cmbFAMILYMember;
            lblFamilyMemberID.Text = HealthClaimApp.Properties.Strings.FamilyMemberID;
            lblEmployeeID.Text = HealthClaimApp.Properties.Strings.employeeID;
            lblRemarks.Text = HealthClaimApp.Properties.Strings.NOTES;
            //Subform
            lblNotes.Text = HealthClaimApp.Properties.Strings.NOTES;
            lblClaimDate.Text = HealthClaimApp.Properties.Strings.lblClaimDATE;
            lblAmountClaimed.Text = HealthClaimApp.Properties.Strings.lblClaimedAmount;
            lblBenefit.Text = HealthClaimApp.Properties.Strings.lblBenefit;
            lblQuantity.Text = HealthClaimApp.Properties.Strings.lblQUANTITY;
            btnSubmit.Text = HealthClaimApp.Properties.Strings.SubmitClaim;
            btnClearBenefit.Text = HealthClaimApp.Properties.Strings.btnCLEARBenefit;
            btnClear.Text = HealthClaimApp.Properties.Strings.btnCLEAR;
            btnAddBenefit.Text = HealthClaimApp.Properties.Strings.BtnAddBenefit;
            btnCancel.Text = HealthClaimApp.Properties.Strings.btnCANCEL;
            //ToolTips 
            ToolTipEmployeeId.SetToolTip(cmbEmployee, HealthClaimApp.Properties.Strings.ToolTipEmployeeID);
            ToolTipFamilyMemberId.SetToolTip(cmbFamilyMember, HealthClaimApp.Properties.Strings.ToolTipFamilyMemberId);
            lblClaimOwner.Text = HealthClaimApp.Properties.Strings.lblClaimOwner;
            ToolTipSwitch.SetToolTip(rjToggleSwitch, HealthClaimApp.Properties.Strings.ToolTipSwitch);
        }
        private void PopulateComboBox(MetroFramework.Controls.MetroComboBox comboBox, string query, string displayMember, string valueMember)
        {
            DataTable table = ExecuteQuery(query);
            if (table != null && table.Rows.Count > 0)
            {
                comboBox.DataSource = table;
                comboBox.DisplayMember = displayMember;
                comboBox.ValueMember = valueMember;
                comboBox.SelectedIndex = -1; // Clear selection
            }
            else
            {
                MetroMessageBox.Show(this, $"No data found for {cmbBenefit.Name}", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public void hideEnableBtns()
        {
            txtDateModified.Visible = false;
            txtDateModified.Enabled = false;
            txtUserModified.Visible = false;
            txtUserModified.Enabled = false;
        }
        private DataTable ExecuteQuery(string query)
        {
            DataTable dataTable = new DataTable();
            using (connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                        {
                            dataAdapter.Fill(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MetroMessageBox.Show(this, $"Error executing query: {ex.Message}", "Query Execution Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    connection.Close();
                }
            }
            return dataTable;
        }

        private void SetupDataGridView()
        {
            dataGridViewBenefits.Columns.Clear();

            dataGridViewBenefits.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "BenefitID",
                HeaderText = "Benefit ID",
                Visible = false
            });
            dataGridViewBenefits.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "BenefitName",
                HeaderText = "Benefit Name",
                ReadOnly = true
            });
            dataGridViewBenefits.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "AmountClaimed",
                HeaderText = "Amount Claimed",
                DefaultCellStyle = { Format = "C2" }
            });
            dataGridViewBenefits.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Quantity",
                HeaderText = "Quantity"
            });
            dataGridViewBenefits.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Notes",
                HeaderText = "Notes"
            });

            dataGridViewBenefits.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridViewBenefits.AllowUserToAddRows = false;
            dataGridViewBenefits.AllowUserToDeleteRows = true;
            dataGridViewBenefits.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            // Update headers after columns are added
            LocalizeDataGridViewBenefitsViewHeaders();
        }

        private void LocalizeDataGridViewBenefitsViewHeaders()
        {
            if (dataGridViewBenefits.Columns.Contains("BenefitName"))
            {
                dataGridViewBenefits.Columns["BenefitName"].HeaderText = HealthClaimApp.Properties.Strings.BenefitNAME;
            }
            if (dataGridViewBenefits.Columns.Contains("AmountClaimed"))
            {
                dataGridViewBenefits.Columns["AmountClaimed"].HeaderText = HealthClaimApp.Properties.Strings.AmountCLAIMED;
            }
            if (dataGridViewBenefits.Columns.Contains("Quantity"))
            {
                dataGridViewBenefits.Columns["Quantity"].HeaderText = HealthClaimApp.Properties.Strings.QUANTITY;
            }
            if (dataGridViewBenefits.Columns.Contains("Notes"))
            {
                dataGridViewBenefits.Columns["Notes"].HeaderText = HealthClaimApp.Properties.Strings.NOTES;
            }
        }


        private void btnAddBenefit_Click(object sender, EventArgs e)
        {
            // Check if a benefit is selected
            int benefitID;
            if (cmbBenefit.SelectedValue == null || !int.TryParse(cmbBenefit.SelectedValue.ToString(), out benefitID))
            {
                MetroMessageBox.Show(this, "Please select a valid benefit.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Get the selected benefit name
            string benefitName = cmbBenefit.Text;

            // Validate and parse the amount claimed
            decimal amountClaimed;
            if (string.IsNullOrWhiteSpace(txtAmountClaimed.Text) ||
                !decimal.TryParse(txtAmountClaimed.Text, out amountClaimed) ||
                amountClaimed <= 0)
            {
                MetroMessageBox.Show(this, "Please enter a valid amount greater than zero.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Validate and parse the quantity
            int quantity;
            if (string.IsNullOrWhiteSpace(txtQuantity.Text) ||
                !int.TryParse(txtQuantity.Text, out quantity) ||
                quantity <= 0)
            {
                MetroMessageBox.Show(this, "Quantity must be a positive integer.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Validate notes input
            string notes = txtNotes.Text.Trim();
            if (string.IsNullOrEmpty(notes))
            {
                MetroMessageBox.Show(this, "Notes cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Create new ClaimDetail instance
            ClaimDetail newDetail = new ClaimDetail(benefitID, benefitName, amountClaimed, quantity, notes);

            // Add new detail to DataGridView
            dataGridViewBenefits.Rows.Add(
                newDetail.BenefitID,
                newDetail.BenefitName,
                newDetail.AmountClaimed,
                newDetail.Quantity,
                newDetail.Notes
            );

            // Clear inputs
            ClearBenefitInputs();
            txtQuantity.Text = "1"; // Reset quantity to default value
        }


        private void ClearBenefitInputs()
        {
            cmbBenefit.SelectedIndex = -1;
            txtAmountClaimed.Clear();
            txtQuantity.Clear();
            txtNotes.Clear();
        }
        //private void PopulateBenefits(string languageCode)
        //{
        //    int activePlanID = GetActivePlanID();
        //    if (activePlanID == 0)
        //    {
        //        MetroMessageBox.Show(this, "No active plan found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return;
        //    }

        //    string benefitQuery = "SELECT BenefitID, BenefitName " +
        //                          "FROM tblBenefits " +
        //                          "WHERE PlanID = ? AND LanguageCode = ? " +
        //                          "ORDER BY BenefitName ASC";
        //    using (OleDbConnection connection = new OleDbConnection(connectionString))
        //    {
        //        try
        //        {
        //            connection.Open();
        //            using (OleDbCommand command = new OleDbCommand(benefitQuery, connection))
        //            {
        //                command.Parameters.AddWithValue("?", activePlanID);
        //                command.Parameters.AddWithValue("?", languageCode);

        //                DataTable table = new DataTable();
        //                using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
        //                {
        //                    adapter.Fill(table);
        //                }

        //                if (table.Rows.Count > 0)
        //                {
        //                    cmbBenefit.DataSource = table;
        //                    cmbBenefit.DisplayMember = "BenefitName";
        //                    cmbBenefit.ValueMember = "BenefitID";
        //                    cmbBenefit.SelectedIndex = -1;
        //                }
        //                else
        //                {
        //                    cmbBenefit.DataSource = null;
        //                    MetroMessageBox.Show(this, "No benefits found for the selected language and active plan. Please contact the administrator if this issue persists.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            MetroMessageBox.Show(this, $"Error loading benefits: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //        finally
        //        {
        //            connection.Close();
        //        }
        //    }
        //}
        private void PopulateBenefits(string languageCode)
        {
            int activePlanID = GetActivePlanID();
            if (activePlanID == 0)
            {
                MetroMessageBox.Show(this, "No active plan found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Updated query to include additional fields
            string benefitQuery = "SELECT BenefitID, BenefitName, Coverage, LimitAmount, Frequency, CoverageType " +
                                  "FROM tblBenefits " +
                                  "WHERE PlanID = ? AND LanguageCode = ? " +
                                  "ORDER BY BenefitName ASC";
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(benefitQuery, connection))
                    {
                        command.Parameters.AddWithValue("?", activePlanID);
                        command.Parameters.AddWithValue("?", languageCode);

                        DataTable table = new DataTable();
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            adapter.Fill(table);
                        }

                        if (table.Rows.Count > 0)
                        {
                            // Format BenefitName to include additional fields
                            foreach (DataRow row in table.Rows)
                            {
                                string formattedName = $"{row["BenefitName"]} - Coverage: {row["Coverage"]}, " +
                                                       $"Limit: {row["LimitAmount"]}, Frequency: {row["Frequency"]}, " +
                                                       $"Type: {row["CoverageType"]}";
                                row["BenefitName"] = formattedName;
                            }

                            cmbBenefit.DataSource = table;
                            cmbBenefit.DisplayMember = "BenefitName";
                            cmbBenefit.ValueMember = "BenefitID";
                            cmbBenefit.SelectedIndex = -1;
                        }
                        else
                        {
                            cmbBenefit.DataSource = null;
                            MetroMessageBox.Show(this, "No benefits found for the selected language and active plan. Please contact the administrator if this issue persists.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MetroMessageBox.Show(this, $"Error loading benefits: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    connection.Close();
                }
            }
        }

        private int GetActivePlanID()
        {
            int activePlanID = 0;
            string query = "SELECT PlanID FROM tblCoveragePlans WHERE IsActive = True";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        object result = command.ExecuteScalar();
                        if (result != null)
                        {
                            activePlanID = Convert.ToInt32(result);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error retrieving active plan: {ex.Message}");
                }
            }

            return activePlanID;
        }
        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate the total claimed amounts against coverage limits
                if (!ValidateClaimAmounts())
                {
                    return; // Exit if validation fails
                }

                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    OleDbTransaction transaction = conn.BeginTransaction();

                    // Prepare the claim insertion query
                    string claimInsertQuery = "INSERT INTO tblClaims (EmployeeID, FamilyMemberID, ClaimDate, TotalAmount, ClaimCategory, Status, Remarks, DateModified, UserModified) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)";
                    using (OleDbCommand cmd = new OleDbCommand(claimInsertQuery, conn, transaction))
                    {
                        int? selectedEmployeeID = (rjToggleSwitch.Checked) ? (int?)cmbEmployee.SelectedValue : null;
                        int? selectedFamilyMemberID = (!rjToggleSwitch.Checked) ? (int?)cmbFamilyMember.SelectedValue : null;

                        // Validate the input based on the toggle switch state
                        if (rjToggleSwitch.Checked && selectedEmployeeID == null)
                        {
                            MetroMessageBox.Show(this, "Please select an employee.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        if (!rjToggleSwitch.Checked && selectedFamilyMemberID == null)
                        {
                            MetroMessageBox.Show(this, "Please select a family member.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        cmd.Parameters.Add("EmployeeID", OleDbType.Integer).Value = selectedEmployeeID ?? (object)DBNull.Value;
                        cmd.Parameters.Add("FamilyMemberID", OleDbType.Integer).Value = selectedFamilyMemberID ?? (object)DBNull.Value;
                        cmd.Parameters.Add("ClaimDate", OleDbType.Date).Value = DateTime.Now;
                        cmd.Parameters.Add("TotalAmount", OleDbType.Currency).Value = CalculateTotalAmount();
                        cmd.Parameters.Add("ClaimCategory", OleDbType.VarChar, 50).Value = rjToggleSwitch.Checked ? "Employee" : "FamilyMember";
                        cmd.Parameters.Add("Status", OleDbType.VarChar, 50).Value = "Pending";
                        cmd.Parameters.Add("Remarks", OleDbType.VarChar, 255).Value = txtRemarks.Text;
                        UserAccount userAccount = new UserAccount();
                        string currentDate = userAccount.GetTimeDate();
                        cmd.Parameters.Add("DateModified", OleDbType.VarChar, 50).Value = currentDate;
                        cmd.Parameters.Add("UserModified", OleDbType.VarChar, 50).Value = UserSession.Instance.Username;

                        cmd.ExecuteNonQuery();
                    }

                    // Retrieve the last inserted claim ID
                    int lastClaimID = GetLastClaimID(conn, transaction);

                    if (lastClaimID > 0)
                    {
                        // Prepare the claim details insertion query
                        string claimDetailInsertQuery = "INSERT INTO tblClaimsDetails (ClaimID, BenefitID, AmountClaimed, Quantity, Notes) VALUES (?, ?, ?, ?, ?)";

                        foreach (DataGridViewRow row in dataGridViewBenefits.Rows)
                        {
                            if (row.IsNewRow) continue;

                            using (OleDbCommand cmd = new OleDbCommand(claimDetailInsertQuery, conn, transaction))
                            {
                                cmd.Parameters.Add("ClaimID", OleDbType.Integer).Value = lastClaimID;
                                cmd.Parameters.Add("BenefitID", OleDbType.Integer).Value = (int)row.Cells["BenefitID"].Value;
                                cmd.Parameters.Add("AmountClaimed", OleDbType.Currency).Value = (decimal)row.Cells["AmountClaimed"].Value;
                                cmd.Parameters.Add("Quantity", OleDbType.Integer).Value = (int)row.Cells["Quantity"].Value;
                                cmd.Parameters.Add("Notes", OleDbType.VarChar, 255).Value =
                                    string.IsNullOrWhiteSpace((string)row.Cells["Notes"].Value) ? (object)DBNull.Value : row.Cells["Notes"].Value;

                                cmd.ExecuteNonQuery();
                            }
                        }

                        // Commit the transaction
                        transaction.Commit();
                        MetroMessageBox.Show(this, "Claim submitted successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ClearForm();
                    }
                    else
                    {
                        MetroMessageBox.Show(this, "Failed to retrieve last claim ID.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MetroMessageBox.Show(this, $"Error submitting claim: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private decimal CalculateTotalAmount()
        {
            decimal totalAmount = 0m;

            foreach (DataGridViewRow row in dataGridViewBenefits.Rows)
            {
                if (row.IsNewRow) continue;

                totalAmount += Convert.ToDecimal(row.Cells["AmountClaimed"].Value);
            }

            return totalAmount;
        }
        private int GetLastClaimID(OleDbConnection conn, OleDbTransaction transaction)
        {
            int lastClaimID = 0;
            string query = "SELECT MAX(ClaimID) FROM tblClaims";

            using (OleDbCommand command = new OleDbCommand(query, conn, transaction))
            {
                object result = command.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    lastClaimID = Convert.ToInt32(result);
                }
            }

            return lastClaimID;
        }

        private void ClearForm()
        {
            cmbEmployee.SelectedIndex = -1;
            cmbFamilyMember.SelectedIndex = -1;
            rjToggleSwitch.Checked = true; // Default to Employee
            txtRemarks.Clear();
            dataGridViewBenefits.Rows.Clear();
        }

        private void rjToggleSwitch1_CheckedChanged(object sender, EventArgs e)
        {
            bool isEmployee = rjToggleSwitch.Checked;
            cmbEmployee.Enabled = isEmployee;
            cmbFamilyMember.Enabled = !isEmployee;
            cmbFamilyMember.SelectedIndex = -1; // Clear if switching to employee
            cmbEmployee.SelectedIndex = -1; // Clear if switching to family member
        }

        // Definition of ClaimDetail class
        public class ClaimDetail
        {
            public int BenefitID { get; }
            public string BenefitName { get; }
            public decimal AmountClaimed { get; }
            public int Quantity { get; }
            public string Notes { get; }

            public ClaimDetail(int benefitID, string benefitName, decimal amountClaimed, int quantity, string notes)
            {
                BenefitID = benefitID;
                BenefitName = benefitName;
                AmountClaimed = amountClaimed;
                Quantity = quantity;
                Notes = notes;
            }
        }
        private void cmbEmployee_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cmbEmployee.SelectedValue != null)
                {
                    // Extract the value from the DataRowView
                    object selectedValue = cmbEmployee.SelectedValue;

                    if (selectedValue is DataRowView)
                    {
                        DataRowView rowView = (DataRowView)selectedValue;
                        // Get the value from the DataRowView
                        selectedValue = rowView["EmployeeID"];
                    }

                    int selectedEmployeeID;
                    if (int.TryParse(selectedValue.ToString(), out selectedEmployeeID))
                    {
                        PopulateFamilyMembers(selectedEmployeeID);

                        // Enable or disable the toggle switch based on the selection
                        rjToggleSwitch.Enabled = true; // Enable toggle switch if employee is selected
                    }
                    else
                    {
                        MessageBox.Show("Selected employee ID is not valid. Please try again.");
                        rjToggleSwitch.Enabled = false; // Disable toggle switch if selection is invalid
                    }
                }
                else
                {
                    // No employee selected
                    cmbFamilyMember.Enabled = false;
                    cmbFamilyMember.SelectedIndex = -1; // Clear selection
                    rjToggleSwitch.Enabled = false; // Disable toggle switch if no employee is selected
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("An error occurred while selecting the employee: {0}", ex.Message));
                rjToggleSwitch.Enabled = false; // Ensure toggle switch is disabled in case of an error
            }
        }

        private void PopulateFamilyMembers(int employeeID)
        {
            string query = "SELECT FamilyMemberID, [FirstName] & ' ' & [LastName] AS FullName " +
                           "FROM tblFamilyMembers WHERE EmployeeID = ? AND IsActive = True ORDER BY [FirstName] & ' ' & [LastName] ASC";

            DataTable table = ExecuteQuery(query, new OleDbParameter("?", employeeID));
            if (table != null && table.Rows.Count > 0)
            {
                cmbFamilyMember.DataSource = table;
                cmbFamilyMember.DisplayMember = "FullName";
                cmbFamilyMember.ValueMember = "FamilyMemberID";
                cmbFamilyMember.SelectedIndex = -1;
            }
            else
            {
                cmbFamilyMember.DataSource = null;
                MessageBox.Show("No family members found for the selected employee.");
            }
        }
        private DataTable ExecuteQuery(string query, params OleDbParameter[] parameters)
        {
            DataTable dataTable = new DataTable();
            using (connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }
                        using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
                        {
                            dataAdapter.Fill(dataTable);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error executing query: {ex.Message}");
                }
                finally
                {
                    connection.Close();
                }
            }
            return dataTable;
        }
        private void rjToggleSwitch_CheckedChanged(object sender, EventArgs e)
        {
            bool isEmployeeClaim = rjToggleSwitch.Checked;

            // Update labels based on the toggle switch state
            UpdateToggleSwitchLabels(isEmployeeClaim);

            if (isEmployeeClaim)
            {
                // When switching to Employee mode
                cmbEmployee.Enabled = true;

                // Disable cmbFamilyMember and clear its selection
                cmbFamilyMember.Enabled = false;
                cmbFamilyMember.SelectedIndex = -1; // Clear selection
            }
            else
            {
                // When switching to Family Member mode
                // Check if there is a selected employee
                if (cmbEmployee.SelectedIndex != -1)
                {
                    // Retain the selected employee
                    // Enable cmbFamilyMember to allow user selection
                    cmbFamilyMember.Enabled = true;
                }
                else
                {
                    // If no employee is selected, disable cmbFamilyMember
                    cmbFamilyMember.Enabled = false;
                }
            }
        }
        private void UpdateToggleSwitchLabels(bool isEmployeeClaim)
        {
            if (isEmployeeClaim)
            {
                lblOn.Text = HealthClaimApp.Properties.Strings.EmployeeClaim;
                lblOff.Text = HealthClaimApp.Properties.Strings.FamilyMemberClaim;
                lblOff.Enabled = false;
                lblOff.Visible = false;
                lblOn.Visible = true;
                lblOn.Enabled = true;
            }
            else
            {
                lblOn.Text = HealthClaimApp.Properties.Strings.EmployeeClaim;
                lblOff.Text = HealthClaimApp.Properties.Strings.FamilyMemberClaim;
                lblOff.Enabled = true;
                lblOff.Visible = true;
                lblOn.Visible = false;
                lblOn.Enabled = false;
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            // Clear text fields
            txtRemarks.Clear();
            cmbEmployee.SelectedIndex = -1;
            cmbFamilyMember.SelectedIndex = -1;
            rjToggleSwitch.Checked = false;
        }
        private void InitializeEmployeeIDSelector()
        {
            cmbEmployee.SelectedIndex = -1;
            cmbEmployee.SelectedIndexChanged += cmbEmployee_SelectedIndexChanged;

            // Set the tooltip text using resx
            ToolTipEmployeeId.SetToolTip(cmbEmployee, HealthClaimApp.Properties.Strings.ToolTipEmployeeID);

            // Initialize the toggle switch state based on employee selection
            rjToggleSwitch.Enabled = false; // Start with toggle switch disabled
        }

        private void InitializeFamilyMemberIDSelector()
        {
            cmbFamilyMember.SelectedIndex = -1;
            cmbFamilyMember.SelectedIndexChanged += cmbFamilyMember_SelectedIndexChanged;

            // Set the tooltip text using resx
            ToolTipFamilyMemberId.SetToolTip(cmbFamilyMember, HealthClaimApp.Properties.Strings.ToolTipFamilyMemberId);
        }

        private void cmbFamilyMember_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Add the logic that should be executed when the selected family member changes.
            try
            {
                if (cmbFamilyMember.SelectedValue != null)
                {
                    // Extract the value from the DataRowView
                    object selectedValue = cmbFamilyMember.SelectedValue;

                    if (selectedValue is DataRowView)
                    {
                        DataRowView rowView = (DataRowView)selectedValue;
                        // Get the value from the DataRowView
                        selectedValue = rowView["FamilyMemberID"];
                    }

                    int selectedFamilyMemberID;
                    if (int.TryParse(selectedValue.ToString(), out selectedFamilyMemberID))
                    {
                        // Logic to handle family member selection
                        // For example, you could populate some related data based on the selected family member
                        // Add your logic here
                    }
                    else
                    {
                        MetroMessageBox.Show(this, "Selected family member ID is not valid. Please try again.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    // No family member selected, you can add any additional logic if needed
                }
            }
            catch (Exception ex)
            {
                MetroMessageBox.Show(this, string.Format("An error occurred while selecting the family member: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private bool ValidateClaimAmounts()
        {
            // Dictionaries to store total claimed amount per benefit ID for employees and family members
            Dictionary<int, decimal> employeeClaimedAmounts = new Dictionary<int, decimal>();
            Dictionary<int, decimal> familyMemberClaimedAmounts = new Dictionary<int, decimal>();

            // Determine if we are working with an employee or a family member
            bool isEmployee = rjToggleSwitch.Checked;
            int? selectedEmployeeID = isEmployee ? (int?)cmbEmployee.SelectedValue : null;
            int? selectedFamilyMemberID = !isEmployee ? (int?)cmbFamilyMember.SelectedValue : null;

            // Iterate over the DataGridView rows to calculate total claimed amounts per benefit ID
            foreach (DataGridViewRow row in dataGridViewBenefits.Rows)
            {
                if (row.IsNewRow) continue;

                int benefitID = Convert.ToInt32(row.Cells["BenefitID"].Value);
                decimal amountClaimed = Convert.ToDecimal(row.Cells["AmountClaimed"].Value);

                if (isEmployee)
                {
                    if (employeeClaimedAmounts.ContainsKey(benefitID))
                    {
                        employeeClaimedAmounts[benefitID] += amountClaimed;
                    }
                    else
                    {
                        employeeClaimedAmounts[benefitID] = amountClaimed;
                    }
                }
                else
                {
                    if (familyMemberClaimedAmounts.ContainsKey(benefitID))
                    {
                        familyMemberClaimedAmounts[benefitID] += amountClaimed;
                    }
                    else
                    {
                        familyMemberClaimedAmounts[benefitID] = amountClaimed;
                    }
                }
            }

            // Validate employee claims
            foreach (var entry in employeeClaimedAmounts)
            {
                int benefitID = entry.Key;
                decimal totalClaimedAmount = entry.Value;

                // Retrieve the coverage limit for the employee benefit
                decimal coverageLimit = GetCoverageLimit(benefitID);

                // Debugging information
                MessageBox.Show($"Employee Benefit ID: {benefitID}, Total Claimed: {totalClaimedAmount}, Coverage Limit: {coverageLimit}");

                // Calculate the total claimed amount for the year for the employee
                decimal totalYearlyClaimedAmount = selectedEmployeeID.HasValue
                    ? GetTotalYearlyClaimedAmount(selectedEmployeeID.Value, benefitID, true)
                    : 0;

                // Check if the total claimed amount exceeds the coverage limit
                if (coverageLimit <= 0)
                {
                    MetroMessageBox.Show(this, $"Coverage limit for employee benefit ID {benefitID} is not valid.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // Check if the total claimed amount exceeds the coverage limit
                if (totalClaimedAmount + totalYearlyClaimedAmount > coverageLimit)
                {
                    MetroMessageBox.Show(this, $"The total claimed amount for employee benefit ID {benefitID} exceeds the coverage limit ({coverageLimit:C}).", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            // Validate family member claims
            foreach (var entry in familyMemberClaimedAmounts)
            {
                int benefitID = entry.Key;
                decimal totalClaimedAmount = entry.Value;

                // Retrieve the coverage limit for the family member benefit
                decimal coverageLimit = GetCoverageLimit(benefitID);

                // Debugging information
                MessageBox.Show($"Family Member Benefit ID: {benefitID}, Total Claimed: {totalClaimedAmount}, Coverage Limit: {coverageLimit}");

                // Calculate the total claimed amount for the year for the family member
                decimal totalYearlyClaimedAmount = selectedFamilyMemberID.HasValue
                    ? GetTotalYearlyClaimedAmount(selectedFamilyMemberID.Value, benefitID, false)
                    : 0;

                // Check if the total claimed amount exceeds the coverage limit
                if (coverageLimit <= 0)
                {
                    MetroMessageBox.Show(this, $"Coverage limit for family member benefit ID {benefitID} is not valid.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                // Check if the total claimed amount exceeds the coverage limit
                if (totalClaimedAmount + totalYearlyClaimedAmount > coverageLimit)
                {
                    MetroMessageBox.Show(this, $"The total claimed amount for family member benefit ID {benefitID} exceeds the coverage limit ({coverageLimit:C}).", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            return true;
        }

        private decimal GetCoverageLimit(int benefitID)
        {
            decimal limitAmount = 0;

            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string query = "SELECT LimitAmount FROM tblBenefits WHERE BenefitID = ?";

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        cmd.Parameters.Add("BenefitID", OleDbType.Integer).Value = benefitID;

                        var result = cmd.ExecuteScalar();

                        if (result != null)
                        {
                            string resultString = result.ToString().Replace("€", "").Replace(",", "."); // Handle currency formatting
                            if (decimal.TryParse(resultString, out limitAmount))
                            {
                                // Successfully parsed limit amount
                            }
                        }

                        // Debugging information
                        MessageBox.Show($"Benefit ID: {benefitID}, Limit Amount: {limitAmount}");
                    }
                }
            }
            catch (Exception ex)
            {
                MetroMessageBox.Show(this, $"Error retrieving coverage limit: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return limitAmount;
        }


        private decimal GetTotalYearlyClaimedAmount(int personID, int benefitID, bool isEmployee)
        {
            decimal totalYearlyClaimedAmount = 0;

            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    DateTime startDate;
                    DateTime endDate;

                    // Retrieve the active coverage period
                    string planQuery = @"
            SELECT StartDate, EndDate
            FROM tblCoveragePlans
            WHERE IsActive = True";

                    using (OleDbCommand planCmd = new OleDbCommand(planQuery, conn))
                    {
                        using (OleDbDataReader reader = planCmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                startDate = reader.GetDateTime(0);
                                endDate = reader.GetDateTime(1);
                            }
                            else
                            {
                                throw new Exception("No active coverage plan found.");
                            }
                        }
                    }

                    // Define the SQL queries
                    string claimQueryEmployee = @"
            SELECT Sum(CD.AmountClaimed) AS TotalClaimed
            FROM tblClaimsDetails AS CD
            INNER JOIN tblClaims AS C ON CD.ClaimID = C.ClaimID
            WHERE C.ClaimDate BETWEEN ? AND ?
            AND CD.BenefitID = ?
            AND C.EmployeeID = ?";

                    string claimQueryFamilyMember = @"
            SELECT Sum(CD.AmountClaimed) AS TotalClaimed
            FROM tblClaimsDetails AS CD
            INNER JOIN tblClaims AS C ON CD.ClaimID = C.ClaimID
            WHERE C.ClaimDate BETWEEN ? AND ?
            AND CD.BenefitID = ?
            AND C.FamilyMemberID = ?";

                    // Use the appropriate query based on isEmployee
                    string claimQuery;
                    if (isEmployee)
                    {
                        claimQuery = claimQueryEmployee;
                    }
                    else
                    {
                        claimQuery = claimQueryFamilyMember;
                    }

                    using (OleDbCommand claimCmd = new OleDbCommand(claimQuery, conn))
                    {
                        claimCmd.Parameters.Add("StartDate", OleDbType.Date).Value = startDate;
                        claimCmd.Parameters.Add("EndDate", OleDbType.Date).Value = endDate;
                        claimCmd.Parameters.Add("BenefitID", OleDbType.Integer).Value = benefitID;
                        claimCmd.Parameters.Add(isEmployee ? "EmployeeID" : "FamilyMemberID", OleDbType.Integer).Value = personID;

                        // Debugging information
                        MessageBox.Show("Executing Query: " + claimQuery + "\nParameters: StartDate=" + startDate + ", EndDate=" + endDate + ", BenefitID=" + benefitID + ", PersonID=" + personID + ", IsEmployee=" + isEmployee);

                        object result = claimCmd.ExecuteScalar();

                        if (result != DBNull.Value && result != null)
                        {
                            string resultString = result.ToString().Replace("€", "").Replace(",", "."); // Handle currency formatting
                            decimal parsedAmount;
                            if (decimal.TryParse(resultString, out parsedAmount))
                            {
                                totalYearlyClaimedAmount = parsedAmount;
                            }
                        }

                        // Debugging information
                        MessageBox.Show("Person ID: " + personID + ", Benefit ID: " + benefitID + ", Total Yearly Claimed Amount: " + totalYearlyClaimedAmount);
                    }
                }
            }
            catch (Exception ex)
            {
                MetroMessageBox.Show(this, "Error retrieving total yearly claimed amount: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return totalYearlyClaimedAmount;
        }


        private void CheckClaimAmount(int personID, int benefitID, bool isEmployee, decimal currentClaimAmount)
        {
            decimal totalClaimedAmount = GetTotalYearlyClaimedAmount(personID, benefitID, isEmployee);
            decimal coverageLimit = GetCoverageLimit(benefitID);

            decimal totalClaimableAmount = coverageLimit - (totalClaimedAmount + currentClaimAmount);

            if (totalClaimableAmount < 0)
            {
                MetroMessageBox.Show(this,
                    $"The total claimed amount for the selected benefit ID exceeds the coverage limit.\n" +
                    $"Current claim: {currentClaimAmount:C}\n" +
                    $"Total claimed this year: {totalClaimedAmount:C}\n" +
                    $"Coverage limit: {coverageLimit:C}\n" +
                    $"Over the limit by: {Math.Abs(totalClaimableAmount):C}",
                    "Validation Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            else
            {
                MetroMessageBox.Show(this,
                    $"The claim amount is acceptable.\n" +
                    $"Total claimed this year: {totalClaimedAmount:C}\n" +
                    $"Remaining balance: {totalClaimableAmount:C}",
                    "Validation Info",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }


        private void btnClearBenefit_Click(object sender, EventArgs e)
        {
            txtNotes.Clear();
            dtpClaimDate.Value = DateTime.Now;
            txtAmountClaimed.Clear();
            txtQuantity.Clear();
            cmbBenefit.SelectedIndex = -1;
        }

        private void txtQuantity_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Check if the pressed key is a digit or control key (like Backspace)
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true; // Ignore the key press
            }
        }
        private void txtAmountClaimed_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Allow control keys (e.g., Backspace, Delete)
            if (char.IsControl(e.KeyChar))
            {
                return;
            }

            // Allow digits
            if (char.IsDigit(e.KeyChar))
            {
                return;
            }

            // Allow only one decimal point
            string decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;
            if (e.KeyChar.ToString() == decimalSeparator && !txtAmountClaimed.Text.Contains(decimalSeparator))
            {
                return;
            }

            // Disallow spaces and any other characters
            e.Handled = true;
        }


        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}