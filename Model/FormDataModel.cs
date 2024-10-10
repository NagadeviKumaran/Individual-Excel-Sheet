using Newtonsoft.Json;

namespace ExcelForm.Model
{
    public class FormDataModel
    {
        

        public string Name { get; set; }
        public string FatherName { get; set; }
        public DateTime DOB { get; set; }
        public string AadharNo { get; set; }
        public string MobileNo { get; set; }
        public string MaritalStatus { get; set; }
        public string Gender { get; set; }
        public string Pan { get; set; }
        public string PresentAddress { get; set; }
        public string PermanentAddress { get; set; }
        public DateTime appointmentDate { get; set; }
        public string dispensaryPreferences { get; set; }
        public string AccountNo { get; set; }
        public string BankName { get; set; }
        public string BranchName { get; set; }
        public string IfscCode { get; set; }


        public List<NomineeModel> Nominees { get; set; } = new List<NomineeModel>();

        public IFormFile nomineeFile { get; set; }
        public List<FamilyDetailModel> FamilyDetails { get; set; } = new List<FamilyDetailModel>();
        public IFormFile familyFile { get; set; }
        public int existingUAN { get; set; }
        public int existingIPN { get; set; }

       
        

        //public string UploadedDocumentPath { get; set; }
    }


    public class NomineeModel
    {
        public string NomineeName { get; set; }
        public string NomineeRelation { get; set; }
        public string NomineeAddress { get; set; }
        public string NomineeAadharNo { get; set; }
        

    }


    public class FamilyDetailModel
    {
        public string FamilyName { get; set; }
        public string FamilyRelation { get; set; }
        public DateTime FamilyDob { get; set; }
        public string FamilyAadharNo { get; set; }
       




    }
}
 


