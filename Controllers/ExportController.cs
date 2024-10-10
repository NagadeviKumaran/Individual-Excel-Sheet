using ExcelForm.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using ClosedXML;
using OfficeOpenXml.Style;
using System.Net.Mail;
using System.Net;
using System.Drawing;
using System.Net.Mime;
using ClosedXML.Excel;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Spreadsheet;
using MailKit.Security;
using MimeKit;
//using ContentType = MimeKit.ContentType;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.EMMA;
using Org.BouncyCastle.Asn1.X509;
using System.Xml.Linq;
using System.Reflection;
using System.Text.Json;
using Aspose.Words.Fields;
using static Org.BouncyCastle.Bcpg.Attr.ImageAttrib;
using System.Data;
using Aspose.Words;
using OfficeOpenXml.DataValidation;

namespace ExcelForm.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExportController : ControllerBase
    {
        [HttpPost("DownloadExcel")]
        public async Task<IActionResult> DownloadExcel([FromForm] FormDataModel formData)
        {


            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;



            // If model state is invalid, return to the form with errors


            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Composite Input Sheet");

                // Header row
                worksheet.Cells["A1:E1"].Merge = true;
                worksheet.Cells["A1"].Value = "Composite Input Sheet for UAN & IPN Allotment [If ESIC Not Applicable Leave S. No 11, 15]";
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Row 2: Fill here
                worksheet.Cells[2, 1].Value = "Fill Here";
                worksheet.Cells[2, 1, 2, 2].Merge = true;

                // Adding input data (dynamic values)
                worksheet.Cells[3, 1].Value = "Name as per Aadhar";
                worksheet.Cells[3, 2].Value = formData.Name;

                worksheet.Cells[4, 1].Value = "Father's Name as per Aadhar";
                worksheet.Cells[4, 2].Value = formData.FatherName;

                worksheet.Cells[5, 1].Value = "Date of Birth as per Aadhar";
                worksheet.Cells[5, 2].Value = formData.DOB.ToShortDateString();

                worksheet.Cells[6, 1].Value = "Aadhar No";
                worksheet.Cells[6, 2].Value = formData.AadharNo;

                worksheet.Cells[7, 1].Value = "Working & Aadhar Linked Mobile No";
                worksheet.Cells[7, 2].Value = formData.MobileNo;

                worksheet.Cells[8, 1].Value = "Martial Status";
                worksheet.Cells[8, 2].Value = formData.MaritalStatus;

                worksheet.Cells[9, 1].Value = "Gender";
                worksheet.Cells[9, 2].Value = formData.Gender;

                worksheet.Cells[10, 1].Value = "PAN";
                worksheet.Cells[10, 2].Value = formData.Pan;

                worksheet.Cells[11, 1].Value = "Present Address";
                worksheet.Cells[11, 2].Value = formData.PresentAddress;
                worksheet.Cells[11, 2, 11, 5].Merge = true;

                worksheet.Cells[12, 1].Value = "Permanent Address";
                worksheet.Cells[12, 2].Value = formData.PermanentAddress;
                worksheet.Cells[12, 2, 12, 5].Merge = true;

                worksheet.Cells[13, 1].Value = "Date of Appointment";
                worksheet.Cells[13, 2].Value = formData.appointmentDate.ToShortDateString();

                worksheet.Cells[14, 1].Value = "Dispensary Preferences [Mention Area]";
                worksheet.Cells[14, 2].Value = formData.dispensaryPreferences;

                // Bank details formatting
                worksheet.Cells[16, 1].Value = "SB Bank Details:";
                worksheet.Cells[16, 1].Style.Font.Bold = true;

                worksheet.Cells[17, 1].Value = "Account No";
                worksheet.Cells[17, 2].Value = formData.AccountNo;

                worksheet.Cells[18, 1].Value = "Bank Name";
                worksheet.Cells[18, 2].Value = formData.BankName;

                worksheet.Cells[19, 1].Value = "Branch Name";
                worksheet.Cells[19, 2].Value = formData.BranchName;

                worksheet.Cells[20, 1].Value = "IFSC Code";
                worksheet.Cells[20, 2].Value = formData.IfscCode;

                // Nominee details
                worksheet.Cells[22, 1].Value = "Nominee Details";
                worksheet.Cells[22, 1, 22, 5].Merge = true;

                worksheet.Cells[23, 1].Value = "Name";
                worksheet.Cells[23, 2].Value = "Relationship";
                worksheet.Cells[23, 3].Value = "Address (if Different)";
                worksheet.Cells[23, 4].Value = "Aadhar No";
               // worksheet.Cells[23, 5].Value = "Nominee File";

                int nomineeRow = 24;
                foreach (var nominee in formData.Nominees)
                {
                    worksheet.Cells[nomineeRow, 1].Value = nominee.NomineeName;
                    worksheet.Cells[nomineeRow, 2].Value = nominee.NomineeRelation;
                    worksheet.Cells[nomineeRow, 3].Value = nominee.NomineeAddress;
                    worksheet.Cells[nomineeRow, 4].Value = nominee.NomineeAadharNo;
                    //worksheet.Cells[nomineeRow, 5].Value = formData.nomineeFile?.FileName;
                    nomineeRow++;
                }

                // Family Members to Add
                worksheet.Cells[nomineeRow + 1, 1].Value = "Family Members to Add";
                worksheet.Cells[nomineeRow + 1, 1, nomineeRow + 1, 5].Merge = true;

                worksheet.Cells[nomineeRow + 2, 1].Value = "Name";
                worksheet.Cells[nomineeRow + 2, 2].Value = "Relationship";
                worksheet.Cells[nomineeRow + 2, 3].Value = "DOB";
                worksheet.Cells[nomineeRow + 2, 4].Value = "Aadhar No";
                

                int familyRow = nomineeRow + 3;
                foreach (var family in formData.FamilyDetails)
                {
                    worksheet.Cells[familyRow, 1].Value = family.FamilyName;
                    worksheet.Cells[familyRow, 2].Value = family.FamilyRelation;
                    worksheet.Cells[familyRow, 3].Value = family.FamilyDob.ToShortDateString();
                    worksheet.Cells[familyRow, 4].Value = family.FamilyAadharNo;
                    familyRow++;
                }

                // Existing UAN and IPN
                worksheet.Cells[familyRow, 1].Value = "Existing UAN (if any)";
                worksheet.Cells[familyRow, 2].Value = formData.existingUAN;

                worksheet.Cells[familyRow + 1, 1].Value = "Existing IPN (if any)";
                worksheet.Cells[familyRow + 1, 2].Value = formData.existingIPN;

                // Final formatting
                worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells.Style.Border.BorderAround(ExcelBorderStyle.Thin);


                // Apply styles
                worksheet.Cells[1, 1, familyRow + 2, 5].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, familyRow + 2, 5].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, familyRow + 2, 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, 1, familyRow + 2, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                worksheet.Cells[1, 1, familyRow + 2, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[1, 1, familyRow + 2, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // Adjust column widths
                worksheet.Column(1).Width = 25;
                worksheet.Column(2).Width = 40;
                worksheet.Column(3).Width = 40;
                worksheet.Column(4).Width = 40;
                worksheet.Column(5).Width = 40;

                var titleCells = new List<string>
                        {
                            "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14", "A16", "A21", "A22","A26","A30","A31"
                        };

                // Apply gray background color to each title cell
                foreach (var cell in titleCells)
                {
                    worksheet.Cells[cell].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[cell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                    worksheet.Cells[cell].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    worksheet.Cells[cell].Style.Font.Bold = true;
                }

                var valueCells = new List<string>
                                    {
                                        "A23","B23","C23","D23","A27","B27","C27","D27"
                                    };


                // Apply light gray background color to each value cell
                foreach (var cell in valueCells)
                {
                    worksheet.Cells[cell].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[cell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    worksheet.Cells[cell].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                }
                var fileContents = package.GetAsByteArray();



                //Convert package to byte array
                    using (var excelStream = new MemoryStream(fileContents))
                {
                   // Prepare to send the email with the attachments
                        var attachments = new List<(Stream stream, string fileName)>
        {
            (excelStream, "FormData.xlsx")
        };

                   // Add nominee document if uploaded
                                        if (formData.nomineeFile != null)
                        {
                            var nomineeStream = formData.nomineeFile.OpenReadStream();
                            attachments.Add((nomineeStream, formData.nomineeFile.FileName));
                        }

                   // Add family document if uploaded
                    if (formData.familyFile != null)
                        {
                            var familyStream = formData.familyFile.OpenReadStream();
                            attachments.Add((familyStream, formData.familyFile.FileName));
                        }

                    //Send email with all attachments
                   var emailSent = await SendEmailWithAttachmentsAsync("dhanushaishu131@gmail.com", "Send the individual WorkSheet", "Test with API", attachments);

                    if (emailSent)
                    {
                        return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FormData.xlsx");
                    }
                    else
                    {
                        return StatusCode(StatusCodes.Status500InternalServerError, "Failed to send email.");
                    }
                }
            }
        }


        // Function to generate Word document for nominees
        private async Task<bool> SendEmailWithAttachmentsAsync(string toEmail, string subject, string body, List<(Stream stream, string fileName)> attachments)
        {
            try
            {
                var smtpClient = new SmtpClient("smtp.gmail.com")
                {
                    Port = 587, // or your SMTP port
                    Credentials = new NetworkCredential("nagadevikumaran01@gmail.com", "lnde fzwr vgbz ybny"),
                    EnableSsl = true,
                };

                var mailMessage = new MailMessage
                {
                    From = new MailAddress("nagadevikumaran01@gmail.com"),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = true,
                };

                mailMessage.To.Add(toEmail);

                // Attach all the provided files
                foreach (var (stream, fileName) in attachments)
                {
                    mailMessage.Attachments.Add(new Attachment(stream, fileName));
                }

                await smtpClient.SendMailAsync(mailMessage);
                return true;
            }
            catch (Exception ex)
            {
                // Log or handle the exception
                return false;
            }
        }
    }
    }





