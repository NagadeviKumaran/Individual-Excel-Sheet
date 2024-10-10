//using ExcelForm.Model;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using OfficeOpenXml.Style;
//using OfficeOpenXml;
//using System.Net.Mail;
//using System.Net;

//namespace ExcelForm.Controllers
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class ExcelFormController : ControllerBase
//    {


//        [HttpPost("UploadDocuments")]
//        public async Task<IActionResult> UploadDocuments([FromForm] IFormFileCollection files)
//        {
//            if (files == null || files.Count == 0)
//            {
//                return BadRequest("No files uploaded.");
//            }

//            // Define the path to store the uploaded documents
//            string uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "UploadedDocuments");

//            // Ensure the directory exists
//            if (!Directory.Exists(uploadPath))
//            {
//                Directory.CreateDirectory(uploadPath);
//            }

//            // List to store file paths
//            var filePaths = new List<string>();

//            foreach (var file in files)
//            {
//                var filePath = Path.Combine(uploadPath, file.FileName);
//                using (var stream = new FileStream(filePath, FileMode.Create))
//                {
//                    await file.CopyToAsync(stream);
//                }
//                filePaths.Add(filePath);
//            }

//            // Return the list of file paths
//            return Ok(new { FilePaths = filePaths });
//        }





//        [HttpPost("DownloadExcel")]
//        public async Task<IActionResult> DownloadExcel([FromBody] FormDataModel formData)
//        {
//            // Set the license context for EPPlus
//            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;




//            using (var package = new ExcelPackage())
//            {
//                var worksheet = package.Workbook.Worksheets.Add("Composite Input Sheet");

//                // Header row
//                worksheet.Cells["A1:E1"].Merge = true;
//                worksheet.Cells["A1"].Value = "Composite Input Sheet for UAN & IPN Allotment [If ESIC Not Applicable Leave S. No 11, 15]";
//                worksheet.Cells["A1"].Style.Font.Bold = true;
//                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

//                // Row 2: Fill here
//                worksheet.Cells[2, 1].Value = "Fill Here";
//                worksheet.Cells[2, 1, 2, 5].Merge = true;

//                // Adding input data (dynamic values)
//                worksheet.Cells[3, 1].Value = "Name as per Aadhar";
//                worksheet.Cells[3, 2].Value = formData.Name;


//                worksheet.Cells[4, 1].Value = "Father's Name as per Aadhar";
//                worksheet.Cells[4, 2].Value = formData.FatherName;

//                worksheet.Cells[5, 1].Value = "Date of Birth as per Aadhar";
//                worksheet.Cells[5, 2].Value = formData.DOB.ToShortDateString();

//                worksheet.Cells[6, 1].Value = "Aadhar No";
//                worksheet.Cells[6, 2].Value = formData.AadharNo;

//                worksheet.Cells[7, 1].Value = "Working & Aadhar Linked Mobile No";
//                worksheet.Cells[7, 2].Value = formData.MobileNo;

//                worksheet.Cells[8, 1].Value = "Martial Status";
//                worksheet.Cells[8, 2].Value = formData.MaritalStatus;

//                worksheet.Cells[9, 1].Value = "Gender";
//                worksheet.Cells[9, 2].Value = formData.Gender;

//                worksheet.Cells[10, 1].Value = "PAN";
//                worksheet.Cells[10, 2].Value = formData.Pan;

//                worksheet.Cells[11, 1].Value = "Present Address";
//                worksheet.Cells[11, 2].Value = formData.PresentAddress;
//                worksheet.Cells[11, 2, 11, 5].Merge = true;

//                worksheet.Cells[12, 1].Value = "Permanent Address";
//                worksheet.Cells[12, 2].Value = formData.PermanentAddress;
//                worksheet.Cells[12, 2, 12, 5].Merge = true;

//                worksheet.Cells[13, 1].Value = "Date of Appointment";
//                worksheet.Cells[13, 2].Value = formData.appointmentDate.ToShortDateString();

//                worksheet.Cells[14, 1].Value = "Dispensary Preferences [Mention Area]";
//                worksheet.Cells[14, 2].Value = formData.dispensaryPreferences;

//                // Bank details on separate rows
//                worksheet.Cells[16, 1].Value = "SB Bank Details:";
//                worksheet.Cells[17, 1].Value = "Account No";
//                worksheet.Cells[17, 2].Value = formData.AccountNo;
//                worksheet.Cells[18, 1].Value = "Bank Name";
//                worksheet.Cells[18, 2].Value = formData.BankName;
//                worksheet.Cells[19, 1].Value = "Branch Name";
//                worksheet.Cells[19, 2].Value = formData.BranchName;
//                worksheet.Cells[20, 1].Value = "IFSC Code";
//                worksheet.Cells[20, 2].Value = formData.IfscCode;

//                // Nominee details
//                worksheet.Cells[22, 1].Value = "Nominee Details";
//                worksheet.Cells[22, 1, 22, 5].Merge = true;

//                worksheet.Cells[23, 1].Value = "Name";
//                worksheet.Cells[23, 2].Value = "Relationship";
//                worksheet.Cells[23, 3].Value = "Address (if Different)";
//                worksheet.Cells[23, 4].Value = "Aadhar No";

//                int nomineeRow = 24;
//                foreach (var nominee in formData.Nominees)
//                {
//                    worksheet.Cells[nomineeRow, 1].Value = nominee.NomineeName;
//                    worksheet.Cells[nomineeRow, 2].Value = nominee.NomineeRelation;
//                    worksheet.Cells[nomineeRow, 3].Value = nominee.NomineeAddress;
//                    worksheet.Cells[nomineeRow, 4].Value = nominee.NomineeAadharNo;
//                    nomineeRow++;
//                }

//                // Family members details
//                worksheet.Cells[nomineeRow + 1, 1].Value = "Family Members to add";
//                worksheet.Cells[nomineeRow + 1, 1, nomineeRow + 1, 5].Merge = true;

//                worksheet.Cells[nomineeRow + 2, 1].Value = "Name";
//                worksheet.Cells[nomineeRow + 2, 2].Value = "Relationship";
//                worksheet.Cells[nomineeRow + 2, 3].Value = "DOB";
//                worksheet.Cells[nomineeRow + 2, 4].Value = "Aadhar No";

//                int familyRow = nomineeRow + 3;
//                foreach (var familyMember in formData.FamilyDetails)
//                {
//                    worksheet.Cells[familyRow, 1].Value = familyMember.FamilyName;
//                    worksheet.Cells[familyRow, 2].Value = familyMember.FamilyRelation;
//                    worksheet.Cells[familyRow, 3].Value = familyMember.FamilyDob.ToShortDateString();
//                    worksheet.Cells[familyRow, 4].Value = familyMember.FamilyAadharNo;
//                    familyRow++;
//                }

//                // Existing UAN and IPN
//                worksheet.Cells[familyRow + 1, 1].Value = "Existing UAN if any";
//                worksheet.Cells[familyRow + 1, 2].Value = formData.existingUAN;

//                worksheet.Cells[familyRow + 2, 1].Value = "Existing IPN if any";
//                worksheet.Cells[familyRow + 2, 2].Value = formData.existingIPN;

//                // Apply styles
//                worksheet.Cells[1, 1, familyRow + 2, 5].Style.Border.Top.Style = ExcelBorderStyle.Thin;
//                worksheet.Cells[1, 1, familyRow + 2, 5].Style.Border.Left.Style = ExcelBorderStyle.Thin;
//                worksheet.Cells[1, 1, familyRow + 2, 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;
//                worksheet.Cells[1, 1, familyRow + 2, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

//                worksheet.Cells[1, 1, familyRow + 2, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
//                worksheet.Cells[1, 1, familyRow + 2, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

//                // Adjust column widths
//                worksheet.Column(1).Width = 25;
//                worksheet.Column(2).Width = 40;
//                worksheet.Column(3).Width = 40;
//                worksheet.Column(4).Width = 40;
//                worksheet.Column(5).Width = 40;

//                var titleCells = new List<string>
//        {
//            "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10", "A11", "A12", "A13", "A14", "A16", "A21", "A22","A26","A30","A31"
//        };

//                // Apply gray background color to each title cell
//                foreach (var cell in titleCells)
//                {
//                    worksheet.Cells[cell].Style.Fill.PatternType = ExcelFillStyle.Solid;
//                    worksheet.Cells[cell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
//                    worksheet.Cells[cell].Style.Font.Color.SetColor(System.Drawing.Color.White);
//                    worksheet.Cells[cell].Style.Font.Bold = true;
//                }

//                var valueCells = new List<string>
//                    {
//                        "A23","B23","C23","D23","A27","B27","C27","D27"
//                    };


//                // Apply light gray background color to each value cell
//                foreach (var cell in valueCells)
//                {
//                    worksheet.Cells[cell].Style.Fill.PatternType = ExcelFillStyle.Solid;
//                    worksheet.Cells[cell].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
//                    worksheet.Cells[cell].Style.Font.Color.SetColor(System.Drawing.Color.Black);
//                }



//                var fileContents = package.GetAsByteArray();
//                var fileName = "Form.xlsx";

//                // Create a memory stream for the generated Excel file
//                using (var excelStream = new MemoryStream(fileContents))
//                {
//                    // Fetch the uploaded document's file path (assuming you store it somewhere or receive it via the formData)
//                    var uploadedDocumentPath = formData.UploadedDocumentPath; // Get the path from form data

//                    if (System.IO.File.Exists(uploadedDocumentPath))
//                    {
//                        // Open the uploaded document as a stream
//                        using (var documentStream = new FileStream(uploadedDocumentPath, FileMode.Open, FileAccess.Read))
//                        {
//                            // Send the email with both attachments
//                            var emailSent = await SendEmailWithAttachmentAsync(

//                                "dhanushaishu131@gmail.com",
//                                "Subject: Excel and Document",
//                                "Body: Here are the files",
//                                excelStream,
//                                "FormData.xlsx",
//                                documentStream,
//                                Path.GetFileName(uploadedDocumentPath)
//                            );

//                            if (emailSent)
//                            {
//                                return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
//                            }
//                            else
//                            {
//                                return StatusCode(StatusCodes.Status500InternalServerError, "Failed to send email.");
//                            }
//                        }
//                    }
//                    else
//                    {
//                        return BadRequest("Uploaded document not found.");
//                    }
//                }
//            }
//        }

//        private async Task<bool> SendEmailWithAttachmentAsync(string toEmail,
//    string subject,
//    string body,
//    Stream excelAttachmentStream,
//    string excelAttachmentName,
//    Stream documentAttachmentStream,
//    string documentAttachmentName)
//        {

//            try
//            {
//                var smtpClient = new SmtpClient("smtp.gmail.com")
//                {
//                    Port = 587,
//                    Credentials = new NetworkCredential("nagadevikumaran01@gmail.com", "lnde fzwr vgbz ybny"),
//                    EnableSsl = true,
//                };

//                var mailMessage = new MailMessage
//                {
//                    From = new MailAddress("nagadevikumaran01@gmail.com"),
//                    Subject = subject,
//                    Body = body,
//                    IsBodyHtml = true,
//                };

//                mailMessage.To.Add(toEmail);

//                // Attach the Excel file
//                mailMessage.Attachments.Add(new Attachment(excelAttachmentStream, excelAttachmentName));

//                // Attach the uploaded document
//                mailMessage.Attachments.Add(new Attachment(documentAttachmentStream, documentAttachmentName));

//                await smtpClient.SendMailAsync(mailMessage);
//                return true;
//            }
//            catch (Exception)
//            {
//                return false;
//            }
//        }




//    }
//}

