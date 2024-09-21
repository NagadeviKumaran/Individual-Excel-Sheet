using ExcelForm.ModelFamilyMemberName;

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

        
        public List<NomineeModel> Nominees { get; set; }
        public List<FamilyDetailModel> FamilyDetails { get; set; }

        public int existingUAN { get; set; }
        public int existingIPN { get; set; }

    }
}
