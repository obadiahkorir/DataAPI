using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CUEReceiver.Controllers
{
    [RoutePrefix("api/values")]
    [EnableCors(origins: "*", headers: "*", methods: "*", SupportsCredentials = true)]
    public class ValuesController : ApiController
    {
        // GET api/values
        //
        [Route("getResearch")]
        public List<Research> GetResearch(String type, String userCode, String password)
        {
            List<Research> allResearch = new List<Research>();
            String universityCode = GetUniversityCode(userCode, password);
            var nav = Config.ReturnNav();
            var researchs = nav.Research.Where(r => r.university == universityCode && r.Category == type);
            foreach (var research in researchs)
            {
                Research myResearch = new Research();
                myResearch.Code = research.Code;
                myResearch.Campus = research.Campus;
                myResearch.Category = research.Category;
                myResearch.Domain = research.Domain;
                myResearch.SubDomain = research.subdomain;
                myResearch.PublicationType = research.Publication_Type;
                myResearch.Title = research.Title;
                myResearch.Description = research.Description;
                myResearch.Link = research.Link;
                myResearch.PatentingOrganisation = research.Patenting_Organisation;
                myResearch.PatentNo = research.Patent_No;
                myResearch.PatentYear = Convert.ToInt32(research.Patent_Year);
                myResearch.AuthorIds = research.Author_Ids;
                allResearch.Add(myResearch);
            }
            return allResearch;
        }

        public String GetUniversityCode(String userCode, String password)
        {
            try
            {
                return new Config().ObjNav().GetInstitutionNumber(userCode, password);
            }
            catch (Exception)
            {

            }
            return "";
        }

        [Route("getStudentsResearch")]
        public List<StudentsResearch> GetStudentsResearch(String type, String userCode, String password)
        {
            List<StudentsResearch> allResearch = new List<StudentsResearch>();
            String universityCode = GetUniversityCode(userCode, password);
            var nav = Config.ReturnNav();
            var researchs = nav.StudentsResearch.Where(r => r.university == universityCode && r.Category == type);
            foreach (var research in researchs)
            {
                StudentsResearch myResearch = new StudentsResearch();
                myResearch.Code = research.Code;
                myResearch.Campus = research.Campus;
                myResearch.Category = research.Category;
                myResearch.Domain = research.Domain;
                myResearch.SubDomain = research.subdomain;
                myResearch.PublicationType = research.Publication_Type;
                myResearch.Title = research.Title;
                myResearch.Description = research.Description;
                myResearch.Link = research.Link;
                myResearch.PatentingOrganisation = research.Patenting_Organisation;
                myResearch.PatentNo = research.Patent_No;
                myResearch.PatentYear = Convert.ToInt32(research.Patent_Year);
                myResearch.AuthorIds = research.Author_Ids;
                allResearch.Add(myResearch);
            }
            return allResearch;
        }

        [Route("getGraduation")]
        public List<Graduation> GetGraduation(String userCode, String password)
        {
            List<Graduation> allGraduation = new List<Graduation>();
            String universityCode = GetUniversityCode(userCode, password);
            var nav = Config.ReturnNav();
            var entries = nav.Graduation.Where(r => r.universityCode == universityCode);
            foreach (var entry in entries)
            {
                Graduation graduation = new Graduation();
                graduation.Id = entry.Id;
                graduation.UniversityCode = entry.universityCode;
                graduation.StudentId = entry.Student_Id;
                graduation.AdmissionNo = entry.Admission_No;
                graduation.FirstName = entry.First_Name;
                graduation.MiddleName = entry.Middle_Name;
                graduation.LastName = entry.Last_Name;
                graduation.ProgramCode = entry.Program_Code;
                graduation.ProgramName = entry.Description;
                graduation.Credit = entry.Credit;
                String date = Convert.ToDateTime(graduation.GraduationDate).ToString("dd-MM-yyyy");
                graduation.GraduationDate = date.Replace("-", "/").Trim();

                allGraduation.Add(graduation);
            }
            return allGraduation;
        }
        [Route("getSubDomains")]
        public String GetSubDomains(String domain)
        {
            var nav = Config.ReturnNav();
            var subdomains = nav.SubDomains.Where(r => r.Domain == domain);
            String response = "";
            foreach (var subdomain in subdomains)
            {
                response += "<option value='" + subdomain.Code + "'>" + subdomain.Description + "</option>";

            }
            return response;
        }

        [Route("getAcademicStaff")]
        public List<AcademicStaff> GetAcademicStaff(String userId, String userPassword)
        {
            List<AcademicStaff> allStaff = new List<AcademicStaff>();
            try
            {
                String universityCode = GetUniversityCode(userId, userPassword);

                var nav = Config.ReturnNav();
                var mstaff = nav.academicstaffentry.Where(r => r.Institution_Code == universityCode);
                foreach (var nStaff in mstaff)
                {
                    AcademicStaff staff = new AcademicStaff();
                    staff.IdNumber = nStaff.ID_Passport_No;
                    staff.SurName = nStaff.Surname;
                    staff.MiddleName = nStaff.Mddle_Name;
                    staff.FirstName = nStaff.First_Name;
                    staff.EthnicBackground = nStaff.Ethnic_Background;
                    String staffDob = Convert.ToDateTime(nStaff.Date_of_Birth).ToString("dd-MM-yyyy");
                    staff.DateOfBirth = staffDob.Replace("-", "/").Trim();
                    staff.HomeCounty = nStaff.Home_County;
                    staff.ProgramDomain = nStaff.Program_Domain;
                    String firstAppointmentDate = Convert.ToDateTime(nStaff.Date_of_first_Appointment).ToString("dd-MM-yyyy");
                    staff.DateOfFirstAppointment = firstAppointmentDate.Replace("-", "/").Trim();
                    staff.Rank = nStaff.Rank;
                    staff.HighestAcademicQualification = nStaff.Highest_Academic_Qualification;
                    staff.PayrollNo = nStaff.Payroll_No;
                    staff.TermsOfService = nStaff.Terms_of_Service;
                    staff.Campus = nStaff.Campus;
                    staff.Gender = nStaff.Gender;
                    staff.Nationality = nStaff.Nationality;
                    staff.DisabilityDescription = nStaff.Disability_Description;
                    staff.DisabilityRegistrationCode = nStaff.Disability_Registration_Code;

                    allStaff.Add(staff);
                }
            }
            catch (Exception u)
            {
                //return u.Message;
            }
            return allStaff;

        }
        [Route("getLibraryStaff")]
        public List<LibraryStaff> GetLibraryStaff(String userId, String userPassword)
        {
            List<LibraryStaff> allStaff = new List<LibraryStaff>();
            try
            {
                String universityCode = GetUniversityCode(userId, userPassword);
                var nav = Config.ReturnNav();
                var mstaff = nav.Library_Staff.Where(r => r.University_Code == universityCode);
                foreach (var nStaff in mstaff)
                {
                    LibraryStaff staff = new LibraryStaff();
                    staff.IdNo = nStaff.Id_Number_Passport_No;
                    staff.FirstName = nStaff.First_Name;
                    staff.MiddleName = nStaff.Middle_Name;
                    staff.LastName = nStaff.Last_Name;
                    staff.Position = nStaff.Position;
                    String staffDob = Convert.ToDateTime(nStaff.Date_of_Birth).ToString("dd-MM-yyyy");
                    staff.DateOfBirth = staffDob.Replace("-", "/").Trim();
                    staff.HighestAcademicQualification = nStaff.Highest_Academic_Qualification;
                    staff.Campus = nStaff.Campus;


                    allStaff.Add(staff);
                }
            }
            catch (Exception u)
            {
                //return u.Message;
            }
            return allStaff;

        }
        [Route("getNonAcademicStaff")]
        public List<NonAcademicStaff> GetNonAcademicStaff(String userId, String userPassword)
        {
            List<NonAcademicStaff> allStaff = new List<NonAcademicStaff>();
            try
            {
                String universityCode = GetUniversityCode(userId, userPassword);
                var nav = Config.ReturnNav();
                var mstaff = nav.GeneralStaff.Where(r => r.University_Code == universityCode);
                foreach (var nStaff in mstaff)
                {
                    NonAcademicStaff staff = new NonAcademicStaff();
                    staff.IDNo = nStaff.ID_Passport_No;
                    staff.Surname = nStaff.Surname;
                    staff.MiddleName = nStaff.Mddle_Name;
                    staff.FirstName = nStaff.First_Name;
                    staff.Gender = nStaff.Gender;
                    staff.EthnicBackground = nStaff.Ethnic_Background;
                    staff.DateOfBirth = Convert.ToDateTime(nStaff.Date_of_Birth).ToString("yyyy-MM-dd");
                    staff.HomeCounty = nStaff.Home_County;
                    staff.Category = nStaff.Category;
                    staff.PayrollNo = nStaff.PayrollNo;
                    staff.Campus = nStaff.Campus;
                    staff.EntryNo = nStaff.Id;
                    allStaff.Add(staff);
                }
            }
            catch (Exception u)
            {
                //return u.Message;
            }
            return allStaff;

        }
        [Route("getPartTimers")]
        public List<PartTimeStaff> GetPartTimers(String userId, String userPassword)
        {
            List<PartTimeStaff> allStaff = new List<PartTimeStaff>();
            try
            {
                String universityCode = GetUniversityCode(userId, userPassword);
                var nav = Config.ReturnNav();
                var mstaff = nav.PartTime.Where(r => r.University_Code == universityCode);
                foreach (var nStaff in mstaff)
                {
                    PartTimeStaff staff = new PartTimeStaff();
                    staff.IDNo = nStaff.ID_Passport_No;
                    staff.Surname = nStaff.Surname;
                    staff.MiddleName = nStaff.Mddle_Name;
                    staff.FirstName = nStaff.First_Name;
                    staff.Gender = nStaff.Gender;
                    staff.EthnicBackground = nStaff.Ethnic_Background;
                    staff.DateOfBirth = Convert.ToDateTime(nStaff.Date_of_Birth).ToString("yyyy-MM-dd");
                    staff.HomeCounty = nStaff.Home_County;
                    staff.stafftype = nStaff.Category;
                    staff.PayrollNo = nStaff.PayrollNo;
                    staff.Campus = nStaff.Campus;
                    staff.EntryNo = nStaff.Id;
                    allStaff.Add(staff);
                }
            }
            catch (Exception u)
            {
                //return u.Message;
            }
            return allStaff;

        }
        [Route("getCases")]
        public List<Discipline> GetCases(String userCode, String password)
        {
            List<Discipline> allCases = new List<Discipline>();
            try
            {
                String universityCode = GetUniversityCode(userCode, password);
                var nav = Config.ReturnNav();
                var cases = nav.DisciplineCases.Where(r => r.University == universityCode);
                foreach (var mCase in cases)
                {
                    Discipline dc = new Discipline();
                    dc.CaseId = mCase.Case_Id + "";
                    dc.StudentId = mCase.Student_Id;
                    dc.StudentName = mCase.Last_Name + " " + mCase.First_Name + " " + mCase.Middle_Name;
                    dc.Description = mCase.Case_Description;
                    dc.CaseDate = Convert.ToDateTime(mCase.Date).ToString("yyyy-MM-dd");
                    dc.Verdict = mCase.Verdict_Category;
                    dc.VerdictDate = Convert.ToDateTime(mCase.Verdict_Date).ToString("yyyy-MM-dd");
                    allCases.Add(dc);
                }
            }
            catch (Exception u)
            {
                //return u.Message;
            }
            return allCases;

        }
        [Route("getAppeals")]
        public List<StudentsAppeals> GetAppeals(String userCode, String password)
        {
            List<StudentsAppeals> allAppeals= new List<StudentsAppeals>();
            try
            {
                String universityCode = GetUniversityCode(userCode, password);
                var nav = Config.ReturnNav();
                var appeals = nav.Appeals.Where(r => r.University == universityCode);
                foreach (var mAppeal in appeals)
                {
                    StudentsAppeals dc = new StudentsAppeals();
                    dc.AppealId = mAppeal.Case_Id + "";
                    dc.StudentId = mAppeal.Student_Id;
                    dc.CaseDescription = mAppeal.Case_Description;
                    dc.StudentName = mAppeal.Last_Name + " " + mAppeal.First_Name + " " + mAppeal.Middle_Name;
                    dc.Verdict = mAppeal.Verdict_Category;
                    dc.VerdictDate = Convert.ToDateTime(mAppeal.Verdict_Date).ToString("yyyy-MM-dd");
                    allAppeals.Add(dc);
                }
            }
            catch (Exception u)
            {
                //return u.Message;
            }
            return allAppeals;

        }
        [Route("getStudents")]
        public List<Student> GetStudents(String userId, String userPassword)
        {
            List<Student> allStudents = new List<Student>();
            try
            {
                String universityCode = GetUniversityCode(userId, userPassword);
                var nav = Config.ReturnNav();
                var students = nav.studentEnrolmentList.Where(r => r.University_Code == universityCode);
                foreach (var student in students)
                {
                    Student myStudent = new Student();
                    myStudent.enrolmentId = Convert.ToInt32(student.Entry_No);
                    myStudent.idNumber = student.ID_Passport_Birth_Certificate;
                    myStudent.Surname = student.Surname;
                    myStudent.firstName = student.First_Name;
                    myStudent.middleName = student.Middle_Name;
                    myStudent.gender = student.Gender;
                    String studentDob = Convert.ToDateTime(student.Date_of_Birth).ToString("dd-MM-yyyy");
                    myStudent.dob = studentDob.Replace("-", "/").Trim();
                    myStudent.homeCounty = student.Home_County;
                    myStudent.ethnicBackground = student.Ethnic_Background;
                    myStudent.nationality = student.Nationality_Code;
                    myStudent.sponsorship = student.Sponsorship;
                    myStudent.disabilityType = student.Disability_Type;
                    myStudent.yearOfStudy = student.Year_of_Study;
                    myStudent.disabilityCode = student.Disability_Code;
                    myStudent.personWithDisability = student.Person_With_Disability.ToString();
                    myStudent.disability = student.Disability;
                    myStudent.ethnicBackground = student.Ethnic_Background;
                    myStudent.myProgram = student.Program;
                    myStudent.programLevel = student.Program_Level;
                    String admissionDate = Convert.ToDateTime(student.Date_of_Admission).ToString("dd-MM-yyyy");
                    myStudent.dateOfAdmission = admissionDate.Replace("-", "/").Trim();
                    myStudent.nationalityName = student.Nationality_Name;
                    myStudent.campus = student.Campus_Code;
                    myStudent.admNo = student.Admission_No;
                    myStudent.stetus = student.Student_Status;
                    myStudent.residents = student.Residential_Status;
                    allStudents.Add(myStudent);
                }
            }
            catch (Exception u)
            {
                //return u.Message;
            }
            return allStudents;

        }

        // GET api/values/5
        public string Get(int id)
        {
            return "value";
        }

        [HttpPost]

        public StudentJson POST([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;

            StudentJson sj = new StudentJson();




            return value; //value2 + "value2";
        }
        [Route("addStudent")]
        public StudentJson AddStudent([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            List<String[]> errorData = new List<string[]>();


            foreach (String[] current in data)
            {
                try
                {
                    int yearOfStudy = 0;
                    Boolean error = false;
                    String message = "";
                    String gender = current[6].Trim();
                    gender = gender.ToLower();
                    int genderCode = 1;
                    if (gender == "male")
                    {
                        genderCode = 0;
                    }
                    else if (gender == "female")
                    {
                        genderCode = 1;
                    }
                    else if (gender == "intersex")
                    {
                        genderCode = 2;
                    }
                    else
                    {
                        error = true;
                        message = "Please enter a valid gender. The only options are (male, female, intersex)";
                    }
                    String stetus = current[18].Trim();
                    int statusCode = 1;
                    try
                    {
                        statusCode = Convert.ToInt32(stetus);
                        if (statusCode > 2 || statusCode < 1)
                        {
                            throw new Exception();
                        }
                        statusCode = statusCode - 1;
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid Students Status. The only options are (Active, Inactive, Transfered,Passed on)";
                    }
                    //else
                    //{
                    //    error = true;
                    //    message = "Please enter a valid Students Status. The only options are (Active, Inactive, Transfered,Passed on)";
                    //}
                    String sponsorship = current[15].Trim();
                    int sponsorshipCode = 1;
                    try
                    {
                        sponsorshipCode = Convert.ToInt32(sponsorship);
                        if (sponsorshipCode >2 || sponsorshipCode < 1)
                        {
                            throw new Exception();
                        }
                        sponsorshipCode = sponsorshipCode - 1;
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid sponsorship code. The only options are (01 for government, 02 for self sponsored)";

                    }
                    String residents = current[17].Trim();
                    int residentsCode = 1;
                    try
                    {
                        residentsCode = Convert.ToInt32(residents);
                        if (residentsCode > 3 || residentsCode < 1)
                        {
                            throw new Exception();
                        }
                        residentsCode = residentsCode - 1;
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid Students Residents code. The only options are (01 for University Accommodation, 02 for Private Accommodation,03 for University Partnership)";

                    }
                    String admNo = String.IsNullOrEmpty(current[1]) ? "" : current[1];
                    if (String.IsNullOrEmpty(current[0]))
                    {
                        error = true;
                        message = "Birth Certificate No/ID No/Passport No is mandatory";
                    }
                    else if ((String.IsNullOrEmpty(current[2]) && String.IsNullOrEmpty(current[3])) || (String.IsNullOrEmpty(current[3]) && String.IsNullOrEmpty(current[4])) || (String.IsNullOrEmpty(current[2]) && String.IsNullOrEmpty(current[4])))
                    {
                        error = true;
                        message = "A student must have at least two names";
                    }
                    DateTime dob = new DateTime();
                    try
                    {
                        dob = DateTime.FromOADate(Convert.ToDouble(current[7]));
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please provide a valid value for date of birth";
                    }
                    try
                    {
                        yearOfStudy = Convert.ToInt32(current[5]);
                        if (yearOfStudy < 1 || yearOfStudy > 10)
                        {
                            throw new Exception();
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please provide a valid value for year of study";
                    }
                    DateTime dateOfAdmission = new DateTime();
                    try
                    {
                        dateOfAdmission = DateTime.FromOADate(Convert.ToDouble(current[12]));
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please provide a valid value for date of admission";
                    }
                    DateTime today = DateTime.Now;
                    if (!error)
                    {
                        if (dateOfAdmission > today)
                        {
                            error = true;
                            message = "Date of admission cannot be earlier than today";
                        }
                        if (dob > today)
                        {
                            error = true;
                            message = "Date of birth cannot be earlier than today";
                        }
                    }
                    if (!error)
                    {
                        //idPassport : Code[30];surname : Text[100];middleName : Text;firstName : Text;gender : Integer;dob : Date;homeCounty : Text;ethnicBckground : Code[10];nationality : Code[10];sponsorship : Integer;
                        //disabilityType : Code[10];yearOfStudy : Text;disabilityCode : Code[10];disability : Text;personWithDisability : Boolean;
                        //myProgram : Code[100];programLevel : Code[10];dateOfAdmission : Date;nationalityName : Text;addedBy : Code[20]) status : Text
                        String tIdNumber = String.IsNullOrEmpty(current[0]) ? "" : current[0];
                        String tSurName = String.IsNullOrEmpty(current[4]) ? "" : current[4];
                        String tMiddleName = String.IsNullOrEmpty(current[3]) ? "" : current[3];
                        String tFirstName = String.IsNullOrEmpty(current[2]) ? "" : current[2];
                        String tEthnicBackground = String.IsNullOrEmpty(current[10]) ? "" : current[10];
                        String nationality = String.IsNullOrEmpty(current[8]) ? "" : current[8];
                        String tHomeCounty = String.IsNullOrEmpty(current[9]) ? "" : current[9];
                        String disabilityCode = String.IsNullOrEmpty(current[14]) ? "" : current[14];
                        String program = String.IsNullOrEmpty(current[11]) ? "" : current[11];
                       

                        String campusCode = "";
                        try
                        {
                            campusCode = String.IsNullOrEmpty(current[16]) ? "" : current[16];
                        }
                        catch (Exception)
                        {
                            campusCode = "";
                        }
                        String disabilityDescription = String.IsNullOrEmpty(current[13]) ? "" : current[13];
                        String status = new Config().ObjNav()
                            .AddStudentDraft(tIdNumber, admNo, tSurName, tMiddleName, tFirstName, genderCode, dob, tHomeCounty,
                                tEthnicBackground, nationality, sponsorshipCode, yearOfStudy, disabilityDescription, disabilityCode,
                              program, dateOfAdmission, campusCode, residentsCode, statusCode, value.userUserName, value.userPassword);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                //convert date of birth and admission date to human readable formats
                                if (i == 7 || i == 12)
                                {
                                    try
                                    {
                                        String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    else
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {
                            errorRecord[i] = current[i];
                            //convert date of birth and date of admission into human readable formats
                            if (i == 7 || i == 12)
                            {
                                try
                                {
                                    String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                    // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                    errorRecord[i] = myDate;
                                }
                                catch (Exception)
                                {
                                    errorRecord[i] = current[i];
                                }
                            }
                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    //value2  = value.column_map;
                }
                catch (Exception t)
                {
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {
                        errorRecord[i] = current[i];
                    }
                    errorRecord[current.Length] = t.Message;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;


            return sj;
        }
        [Route("addPublication")]
        public StudentJson AddPublication([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            List<String[]> errorData = new List<string[]>();
            foreach (String[] current in data)
            {
                try
                {
                    Boolean error = false;
                    String message = "";
                    String domain = "";
                    try
                    {
                        domain = current[0].Trim();
                        if (domain.Length < 1)
                        {
                            throw new Exception();
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please specify the domain of this publication";
                    }
                    String subDomain = "";//1
                    try
                    {
                        subDomain = current[1].Trim();
                    }
                    catch (Exception)
                    {
                        subDomain = "";
                    }
                    String campus = "";//2
                    try
                    {
                        campus = current[2].Trim();
                    }
                    catch (Exception)
                    {
                        campus = "";
                    }
                    String publicationType = "";//3
                    int publicationTypeCode = 0;
                    try
                    {
                        publicationType = current[3].Trim();
                        publicationType = publicationType.ToLower();
                        if (publicationType == "book")
                        {
                            publicationTypeCode = 0;
                        }
                        else if (publicationType == "journal")
                        {
                            publicationTypeCode = 1;
                        }
                        else if (publicationType == "audio-visual")
                        {
                            publicationTypeCode = 2;
                        }
                        else
                        {
                        error = true;
                        message += "Please provide a valid value for publication type. The only possible options are book, journal and audiovisual";
                        }
                    }
                    catch (Exception)
                    {
                        publicationType = "";
                    }
                    String title = "";//4
                    try
                    {
                        title = current[4].Trim();
                    }
                    catch (Exception)
                    {
                        title = "";
                    }
                    String description = "";//5
                    try
                    {
                        description = current[5].Trim();
                    }
                    catch (Exception)
                    {
                        description = "";
                    }
                    String link = "";//6
                    try
                    {
                        link = current[6].Trim();
                    }
                    catch (Exception)
                    {
                        link = "";
                    }
                    String publisher = "";//7
                    try
                    {
                        publisher = current[7].Trim();
                    }
                    catch (Exception)
                    {
                        publisher = "";
                    }
                    String doiNumber = "";//8
                    try
                    {
                        doiNumber = current[8].Trim();
                    }
                    catch (Exception)
                    {
                        doiNumber = "";
                    }
                    int yearPublished = 0;//9
                    try
                    {
                        if (current[9].Trim().Length < 1)
                        {
                            error = true;
                            message += "Please specify the year the publication was published";
                        }
                        else
                        {
                            yearPublished = Convert.ToInt32(current[9].Trim());
                        }
                    }
                    catch (Exception)
                    {
                        yearPublished = 0;
                    }
                    String authorIds = "";//10
                    try
                    {
                        authorIds = current[10].Trim();
                    }
                    catch (Exception)
                    {
                        authorIds = "";
                    }


                    if (!error)
                    {
                        String status = new Config().ObjNav()
                            .AddResearch(0, domain, subDomain, campus, publicationTypeCode, title, description, link,
                                       publisher, doiNumber, yearPublished, authorIds, value.userUserName, value.userPassword);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                //convert date of birth and admission date to human readable formats
                                if (i == 7 || i == 12)
                                {
                                    try
                                    {
                                        String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    else
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {
                            errorRecord[i] = current[i];
                            //convert date of birth and date of admission into human readable formats

                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    //value2  = value.column_map;
                }
                catch (Exception t)
                {
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {
                        errorRecord[i] = current[i];
                    }
                    errorRecord[current.Length] = t.Message;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;


            return sj;
        }
        [Route("addStudentsPublication")]
        public StudentJson AddStudentsPublication([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            List<String[]> errorData = new List<string[]>();


            foreach (String[] current in data)
            {
                try
                {

                    Boolean error = false;
                    String message = "";
                    String domain = "";
                    try
                    {
                        domain = current[0].Trim();
                        if (domain.Length < 1)
                        {
                            throw new Exception();
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please specify the domain of this publication";
                    }
                    String subDomain = "";//1
                    try
                    {
                        subDomain = current[1].Trim();
                    }
                    catch (Exception)
                    {
                        subDomain = "";
                    }
                    String campus = "";//2
                    try
                    {
                        campus = current[2].Trim();
                    }
                    catch (Exception)
                    {
                        campus = "";
                    }
                    String publicationType = "";//3
                    int publicationTypeCode = 0;
                    try
                    {
                        publicationType = current[3].Trim();
                        publicationType = publicationType.ToLower();
                        if (publicationType == "book")
                        {
                            publicationTypeCode = 0;
                        }
                        else if (publicationType == "journal")
                        {
                            publicationTypeCode = 1;
                        }
                        else if (publicationType == "audio-visual")
                        {
                            publicationTypeCode = 2;
                        }
                        else
                        {
                            error = true;
                            message += "Please provide a valid value for publication type. The only possible options are book, journal and audiovisual";
                        }
                    }
                    catch (Exception)
                    {
                        publicationType = "";
                    }
                    String title = "";//4
                    try
                    {
                        title = current[4].Trim();
                    }
                    catch (Exception)
                    {
                        title = "";
                    }
                    String description = "";//5
                    try
                    {
                        description = current[5].Trim();
                    }
                    catch (Exception)
                    {
                        description = "";
                    }
                    String link = "";//6
                    try
                    {
                        link = current[6].Trim();
                    }
                    catch (Exception)
                    {
                        link = "";
                    }
                    String publisher = "";//7
                    try
                    {
                        publisher = current[7].Trim();
                    }
                    catch (Exception)
                    {
                        publisher = "";
                    }
                    String doiNumber = "";//8
                    try
                    {
                        doiNumber = current[8].Trim();
                    }
                    catch (Exception)
                    {
                        doiNumber = "";
                    }
                    int yearPublished = 0;//9
                    try
                    {
                        if (current[9].Trim().Length < 1)
                        {
                            error = true;
                            message += "Please specify the year the publication was published";
                        }
                        else
                        {
                            yearPublished = Convert.ToInt32(current[9].Trim());
                        }
                    }
                    catch (Exception)
                    {
                        yearPublished = 0;
                    }
                    String authorIds = "";//10
                    try
                    {
                        authorIds = current[10].Trim();
                    }
                    catch (Exception)
                    {
                        authorIds = "";
                    }


                    if (!error)
                    {
                        String status = new Config().ObjNav()
                            .AddStudentsResearch(0, domain, subDomain, campus, publicationTypeCode, title, description, link,
                    publisher, doiNumber, yearPublished, authorIds, value.userUserName, value.userPassword);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                //convert date of birth and admission date to human readable formats
                                if (i == 7 || i == 12)
                                {
                                    try
                                    {
                                        String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    else
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {
                            errorRecord[i] = current[i];
                            //convert date of birth and date of admission into human readable formats

                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    //value2  = value.column_map;
                }
                catch (Exception t)
                {
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {
                        errorRecord[i] = current[i];
                    }
                    errorRecord[current.Length] = t.Message;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;


            return sj;
        
    }


        [Route("addPatent")]
        public StudentJson AddPatent([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            List<String[]> errorData = new List<string[]>();


            foreach (String[] current in data)
            {
                try
                {

                    Boolean error = false;
                    String message = "";
                    String domain = "";
                    try
                    {
                        domain = current[0].Trim();
                        if (domain.Length < 1)
                        {
                            throw new Exception();
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please specify the domain of this publication";
                    }
                    String subDomain = "";//1
                    try
                    {
                        subDomain = current[1].Trim();
                    }
                    catch (Exception)
                    {
                        subDomain = "";
                    }
                    String campus = "";//2
                    try
                    {
                        campus = current[2].Trim();
                    }
                    catch (Exception)
                    {
                        campus = "";
                    }

                    int publicationTypeCode = 0;

                    String title = "";//3
                    try
                    {
                        title = current[3].Trim();
                    }
                    catch (Exception)
                    {
                        title = "";
                    }
                    String description = "";//4
                    try
                    {
                        description = current[4].Trim();
                    }
                    catch (Exception)
                    {
                        description = "";
                    }
                    String link = "";//5
                    try
                    {
                        link = current[5].Trim();
                    }
                    catch (Exception)
                    {
                        link = "";
                    }
                    String publisher = "";//6
                    try
                    {
                        publisher = current[6].Trim();
                    }
                    catch (Exception)
                    {
                        publisher = "";
                    }
                    String doiNumber = "";//7
                    try
                    {
                        doiNumber = current[7].Trim();
                    }
                    catch (Exception)
                    {
                        doiNumber = "";
                    }
                    int yearPublished = 0;//8
                    try
                    {
                        if (current[8].Trim().Length < 1)
                        {
                            error = true;
                            message += "Please specify the year the publication was published";
                        }
                        else
                        {
                            yearPublished = Convert.ToInt32(current[8].Trim());
                        }
                    }
                    catch (Exception)
                    {
                        yearPublished = 0;
                    }
                    String authorIds = "";//9
                    try
                    {
                        authorIds = current[9].Trim();
                    }
                    catch (Exception)
                    {
                        authorIds = "";
                    }


                    if (!error)
                    {
                        String status = new Config().ObjNav()
                            .AddResearch(2, domain, subDomain, campus, publicationTypeCode, title, description, link,
                    publisher, doiNumber, yearPublished, authorIds, value.userUserName, value.userPassword);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                //convert date of birth and admission date to human readable formats
                                if (i == 7 || i == 12)
                                {
                                    try
                                    {
                                        String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    else
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {
                            errorRecord[i] = current[i];
                            //convert date of birth and date of admission into human readable formats

                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    //value2  = value.column_map;
                }
                catch (Exception t)
                {
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {
                        errorRecord[i] = current[i];
                    }
                    errorRecord[current.Length] = t.Message;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;


            return sj;
        }
        [Route("addStudentsPatent")]
        public StudentJson AddStudentsPatent([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            List<String[]> errorData = new List<string[]>();


            foreach (String[] current in data)
            {
                try
                {

                    Boolean error = false;
                    String message = "";
                    String domain = "";
                    try
                    {
                        domain = current[0].Trim();
                        if (domain.Length < 1)
                        {
                            throw new Exception();
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please specify the domain of this publication";
                    }
                    String subDomain = "";//1
                    try
                    {
                        subDomain = current[1].Trim();
                    }
                    catch (Exception)
                    {
                        subDomain = "";
                    }
                    String campus = "";//2
                    try
                    {
                        campus = current[2].Trim();
                    }
                    catch (Exception)
                    {
                        campus = "";
                    }

                    int publicationTypeCode = 0;//3

                    String title = "";//4
                    try
                    {
                        title = current[4].Trim();
                    }
                    catch (Exception)
                    {
                        title = "";
                    }
                    String description = "";//5
                    try
                    {
                        description = current[5].Trim();
                    }
                    catch (Exception)
                    {
                        description = "";
                    }
                    String link = "";//6
                    try
                    {
                        link = current[6].Trim();
                    }
                    catch (Exception)
                    {
                        link = "";
                    }
                    String publisher = "";//7
                    try
                    {
                        publisher = current[7].Trim();
                    }
                    catch (Exception)
                    {
                        publisher = "";
                    }
                    String doiNumber = "";//8
                    try
                    {
                        doiNumber = current[8].Trim();
                    }
                    catch (Exception)
                    {
                        doiNumber = "";
                    }
                    int yearPublished = 0;//9
                    try
                    {
                        if (current[9].Trim().Length < 1)
                        {
                            error = true;
                            message += "Please specify the year the publication was published";
                        }
                        else
                        {
                            yearPublished = Convert.ToInt32(current[9].Trim());
                        }
                    }
                    catch (Exception)
                    {
                        yearPublished = 0;
                    }
                    String authorIds = "";//10
                    try
                    {
                        authorIds = current[10].Trim();
                    }
                    catch (Exception)
                    {
                        authorIds = "";
                    }


                    if (!error)
                    {
                        String status = new Config().ObjNav()
                          .AddStudentsResearch(2, domain, subDomain, campus, publicationTypeCode, title, description, link,
                    publisher, doiNumber, yearPublished, authorIds, value.userUserName, value.userPassword);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                //convert date of birth and admission date to human readable formats
                                if (i == 7 || i == 12)
                                {
                                    try
                                    {
                                        String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    else
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {
                            errorRecord[i] = current[i];
                            //convert date of birth and date of admission into human readable formats

                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    //value2  = value.column_map;
                }
                catch (Exception t)
                {
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {
                        errorRecord[i] = current[i];
                    }
                    errorRecord[current.Length] = t.Message;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;


            return sj;
        }
        [Route("addInnovation")]
        public StudentJson AddInnovation([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            List<String[]> errorData = new List<string[]>();


            foreach (String[] current in data)
            {
                try
                {

                    Boolean error = false;
                    String message = "";
                    String domain = "";
                    try
                    {
                        domain = current[0].Trim();
                        if (domain.Length < 1)
                        {
                            throw new Exception();
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please specify the domain of this publication";
                    }
                    String subDomain = "";//1
                    try
                    {
                        subDomain = current[1].Trim();
                    }
                    catch (Exception)
                    {
                        subDomain = "";
                    }
                    String campus = "";//2
                    try
                    {
                        campus = current[2].Trim();
                    }
                    catch (Exception)
                    {
                        campus = "";
                    }

                    int publicationTypeCode = 0;//3

                    String title = "";//4
                    try
                    {
                        title = current[4].Trim();
                    }
                    catch (Exception)
                    {
                        title = "";
                    }
                    String description = "";//5
                    try
                    {
                        description = current[5].Trim();
                    }
                    catch (Exception)
                    {
                        description = "";
                    }
                    String link = "";//6
                    try
                    {
                        link = current[6].Trim();
                    }
                    catch (Exception)
                    {
                        link = "";
                    }
                    String publisher = "";//7
                    try
                    {
                        publisher = current[7].Trim();
                    }
                    catch (Exception)
                    {
                        publisher = "";
                    }
                    String doiNumber = "";//8
                    try
                    {
                        doiNumber = current[8].Trim();
                    }
                    catch (Exception)
                    {
                        doiNumber = "";
                    }
                    int yearPublished = 0;//9
                    try
                    {
                        if (current[9].Trim().Length < 1)
                        {
                            error = true;
                            message += "Please specify the year the publication was published";
                        }
                        else
                        {
                            yearPublished = Convert.ToInt32(current[9].Trim());
                        }
                    }
                    catch (Exception)
                    {
                        yearPublished = 0;
                    }
                    String authorIds = "";//10
                    try
                    {
                        authorIds = current[10].Trim();
                    }
                    catch (Exception)
                    {
                        authorIds = "";
                    }


                    if (!error)
                    {
                        String status = new Config().ObjNav()
                            .AddResearch(1, domain, subDomain, campus, publicationTypeCode, title, description, link,
                    publisher, doiNumber, yearPublished, authorIds, value.userUserName, value.userPassword);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                //convert date of birth and admission date to human readable formats
                                if (i == 7 || i == 12)
                                {
                                    try
                                    {
                                        String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    else
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {
                            errorRecord[i] = current[i];
                            //convert date of birth and date of admission into human readable formats

                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    //value2  = value.column_map;
                }
                catch (Exception t)
                {
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {
                        errorRecord[i] = current[i];
                    }
                    errorRecord[current.Length] = t.Message;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;


            return sj;
        }
        [Route("addStudentsInnovation")]
        public StudentJson AddStudentsInnovation([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            List<String[]> errorData = new List<string[]>();


            foreach (String[] current in data)
            {
                try
                {

                    Boolean error = false;
                    String message = "";
                    String domain = "";
                    try
                    {
                        domain = current[0].Trim();
                        if (domain.Length < 1)
                        {
                            throw new Exception();
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please specify the domain of this publication";
                    }
                    String subDomain = "";//1
                    try
                    {
                        subDomain = current[1].Trim();
                    }
                    catch (Exception)
                    {
                        subDomain = "";
                    }
                    String campus = "";//2
                    try
                    {
                        campus = current[2].Trim();
                    }
                    catch (Exception)
                    {
                        campus = "";
                    }

                    int publicationTypeCode = 0;//3

                    String title = "";//4
                    try
                    {
                        title = current[4].Trim();
                    }
                    catch (Exception)
                    {
                        title = "";
                    }
                    String description = "";//5
                    try
                    {
                        description = current[5].Trim();
                    }
                    catch (Exception)
                    {
                        description = "";
                    }
                    String link = "";//6
                    try
                    {
                        link = current[6].Trim();
                    }
                    catch (Exception)
                    {
                        link = "";
                    }
                    String publisher = "";//7
                    try
                    {
                        publisher = current[7].Trim();
                    }
                    catch (Exception)
                    {
                        publisher = "";
                    }
                    String doiNumber = "";//8
                    try
                    {
                        doiNumber = current[8].Trim();
                    }
                    catch (Exception)
                    {
                        doiNumber = "";
                    }
                    int yearPublished = 0;//9
                    try
                    {
                        if (current[9].Trim().Length < 1)
                        {
                            error = true;
                            message += "Please specify the year the publication was published";
                        }
                        else
                        {
                            yearPublished = Convert.ToInt32(current[9].Trim());
                        }
                    }
                    catch (Exception)
                    {
                        yearPublished = 0;
                    }
                    String authorIds = "";//10
                    try
                    {
                        authorIds = current[10].Trim();
                    }
                    catch (Exception)
                    {
                        authorIds = "";
                    }


                    if (!error)
                    {
                        String status = new Config().ObjNav()
                             .AddStudentsResearch(1, domain, subDomain, campus, publicationTypeCode, title, description, link,
                       publisher, doiNumber, yearPublished, authorIds, value.userUserName, value.userPassword);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                //convert date of birth and admission date to human readable formats
                                if (i == 7 || i == 12)
                                {
                                    try
                                    {
                                        String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    else
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {
                            errorRecord[i] = current[i];
                            //convert date of birth and date of admission into human readable formats

                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    //value2  = value.column_map;
                }
                catch (Exception t)
                {
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {
                        errorRecord[i] = current[i];
                    }
                    errorRecord[current.Length] = t.Message;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;


            return sj;
        }
        [Route("addGraduation")]
        public StudentJson AddGraduation([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            List<String[]> errorData = new List<string[]>();


            foreach (String[] current in data)
            {
                try
                {

                    Boolean error = false;
                    String message = "";
                    String idNo = "";// 0
                    try
                    {
                        idNo = current[0].Trim();
                    }
                    catch (Exception)
                    {
                        idNo = "";
                    }
                    String admissionNo = "";//1
                    try
                    {
                        admissionNo = current[1].Trim();
                    }
                    catch (Exception)
                    {
                        admissionNo = "";
                    }
                    String firstName = "";//2
                    try
                    {
                        firstName = current[2].Trim();
                    }
                    catch (Exception)
                    {
                        firstName = "";
                    }
                    String middleName = "";//3
                    try
                    {
                        middleName = current[3].Trim();
                    }
                    catch (Exception)
                    {
                        middleName = "";
                    }
                    String lastName = "";//4
                    try
                    {
                        lastName = current[4].Trim();
                    }
                    catch (Exception)
                    {
                        lastName = "";
                    }
                    String programCode = "";//5
                    try
                    {
                        programCode = current[5].Trim();
                    }
                    catch (Exception)
                    {
                        programCode = "";
                    }
                    String credit = "";//6
                    try
                    {
                        credit = current[6].Trim();
                    }
                    catch (Exception)
                    {
                        credit = "";
                    }

                    DateTime graduationDate = new DateTime();
                    try
                    {
                        graduationDate = DateTime.FromOADate(Convert.ToDouble(current[7]));
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please provide a valid value graduation date";
                    }
                    if (String.IsNullOrEmpty(idNo))
                    {
                        error = true;
                        message = "Birth Certificate No/ID No/Passport No is mandatory";
                    }
                    else if ((String.IsNullOrEmpty(current[2]) && String.IsNullOrEmpty(current[3])) || (String.IsNullOrEmpty(current[3]) && String.IsNullOrEmpty(current[4])) || (String.IsNullOrEmpty(current[2]) && String.IsNullOrEmpty(current[4])))
                    {
                        error = true;
                        message = "A student must have at least two names";
                    }
                    else if (String.IsNullOrEmpty(programCode))
                    {
                        error = true;
                        message = "Program Code is mandatory";
                    }
                    if (!error)
                    {
                        String status = new Config().ObjNav()
                            .AddGraduation(idNo, admissionNo, firstName, middleName, lastName, programCode, credit, graduationDate, value.userUserName, value.userPassword);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                //convert date of birth and admission date to human readable formats
                                if (i == 7)
                                {
                                    try
                                    {
                                        String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    else
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {
                            errorRecord[i] = current[i];
                            if (i == 7)
                            {
                                try
                                {
                                    String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                    // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                    errorRecord[i] = myDate;
                                }
                                catch (Exception)
                                {
                                    errorRecord[i] = current[i];
                                }
                            }

                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    //value2  = value.column_map;
                }
                catch (Exception t)
                {
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {
                        errorRecord[i] = current[i];
                        if (i == 7)
                        {
                            try
                            {
                                String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                errorRecord[i] = myDate;
                            }
                            catch (Exception)
                            {
                                errorRecord[i] = current[i];
                            }
                        }
                    }
                    errorRecord[current.Length] = t.Message;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;


            return sj;
        }

        [Route("addAcademicStaff")]
        public StudentJson AddAcademicStaff([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;


            List<String[]> errorData = new List<string[]>();
            foreach (String[] current in data)
            {
                try
                {
                    String message = "";
                    Boolean error = false;
                    String idNumber = "";
                    try
                    {
                        idNumber = String.IsNullOrEmpty(current[0].Trim()) ? "" : current[0].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String payrollNo = "";
                    try
                    {
                        payrollNo = String.IsNullOrEmpty(current[1].Trim()) ? "" : current[1].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String firstName = "";
                    try
                    {
                        firstName = String.IsNullOrEmpty(current[2].Trim()) ? "" : current[2].Trim();
                    }
                    catch (Exception E)
                    {


                    }
                    String middleName = "";
                    try
                    {
                        middleName = String.IsNullOrEmpty(current[3].Trim()) ? "" : current[3].Trim();
                    }
                    catch (Exception) { }
                    String lastName = "";
                    try
                    {
                        lastName = String.IsNullOrEmpty(current[4].Trim()) ? "" : current[4].Trim();
                    }
                    catch (Exception) { }
                    if (String.IsNullOrEmpty(idNumber))
                    {
                        error = true;
                        message = "ID No/Passport No is mandatory";
                    }
                    else if ((String.IsNullOrEmpty(firstName) && String.IsNullOrEmpty(middleName)) ||
                             (String.IsNullOrEmpty(middleName) && String.IsNullOrEmpty(lastName)) ||
                             (String.IsNullOrEmpty(lastName) && String.IsNullOrEmpty(firstName)))
                    {
                        error = true;
                        message = "A staff must have at least two names";
                    }
                    String gender = "";
                    try
                    {
                        gender = String.IsNullOrEmpty(current[5].Trim()) ? "" : current[5].Trim();
                    }
                    catch (Exception) { }
                    gender = gender.ToLower();
                    int genderCode = 1;
                    if (gender == "male")
                    {
                        genderCode = 0;
                    }
                    else if (gender == "female")
                    {
                        genderCode = 1;
                    }
                    else if (gender == "intersex")
                    {
                        genderCode = 2;
                    }
                    else
                    {
                        error = true;
                        message = "Please enter a valid gender. The only options are (male, female, intersex)";
                    }
                    String ethnicBackground = "";
                    try
                    {
                        ethnicBackground = String.IsNullOrEmpty(current[6].Trim()) ? "" : current[6].Trim();
                    }
                    catch (Exception) { }
                    String myDob = "";
                    try
                    {
                        myDob = String.IsNullOrEmpty(current[7].Trim()) ? "" : current[7].Trim();
                    }
                    catch (Exception) { }
                    DateTime dateOfBirth = new DateTime();
                    try
                    {
                        dateOfBirth = DateTime.FromOADate(Convert.ToDouble(myDob));
                        if (dateOfBirth > DateTime.Now)
                        {
                            error = true;
                            message = "Date of birth cannot be earlier than today";
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid value for date of birth";
                    }
                    String nationalityCode = "";
                    try
                    {
                        nationalityCode = String.IsNullOrEmpty(current[8].Trim()) ? "" : current[8].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String homeCounty = "";
                    try
                    {
                        homeCounty = String.IsNullOrEmpty(current[9].Trim()) ? "" : current[9].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String disabilityDescription = "";
                    try
                    {
                        disabilityDescription = String.IsNullOrEmpty(current[10].Trim()) ? "" : current[10].Trim();
                    }
                    catch (Exception) { }
                    String disabilityRegistrationCode = "";
                    try
                    {
                        disabilityRegistrationCode = String.IsNullOrEmpty(current[11].Trim())
                        ? ""
                        : current[11].Trim();
                    }
                    catch (Exception) {
                    }
                    String domainCode = "";
                    try
                    {
                        domainCode = String.IsNullOrEmpty(current[12].Trim()) ? "" : current[12].Trim();
                    }
                    catch (Exception) {
                    }

                    String rankCode = "";
                    try
                    {
                        rankCode = String.IsNullOrEmpty(current[13].Trim()) ? "" : current[13].Trim();
                    }
                    catch (Exception) {
                        error = true;
                        message = "Please enter a valid value Staff rank";
                    }
                    String appointmentDate = "";
                    try
                    {
                        appointmentDate = String.IsNullOrEmpty(current[14].Trim()) ? "" : current[14].Trim();
                    }
                    catch (Exception) {
                     }
                    DateTime dateOfFirstAppointment = new DateTime();
                    try
                    {
                        dateOfFirstAppointment = DateTime.FromOADate(Convert.ToDouble(appointmentDate));
                        if (dateOfFirstAppointment > DateTime.Now)
                        {
                            error = true;
                            message = "Date of first appointment cannot be earlier than today";
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid value for date of first appointment";
                    }
                     String currentappointmentDate = "";
                    try
                    {
                        currentappointmentDate = String.IsNullOrEmpty(current[18].Trim()) ? "" : current[18].Trim();
                    }
                    catch (Exception) {
                     }
                    DateTime dateOfCurrentAppointment = new DateTime();
                    try
                    {
                        dateOfCurrentAppointment = DateTime.FromOADate(Convert.ToDouble(currentappointmentDate));
                        if (dateOfCurrentAppointment > DateTime.Now)
                        {
                            error = true;
                            message = "Date of  Current appointment cannot be earlier than today";
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid value for Date of Current appointment";
                    }
                    String termsOfService = "";
                    try
                    {
                        termsOfService = String.IsNullOrEmpty(current[15].Trim()) ? "" : current[15].Trim();
                    }
                    catch (Exception) {
                    }
                    int termsOfServiceCode = 0;
                    try
                    {
                        termsOfServiceCode = Convert.ToInt32(termsOfService);
                        if (termsOfServiceCode > 2 || termsOfServiceCode < 1)
                        {
                            throw new Exception();
                        }
                        termsOfServiceCode -= 1;
                    }
                    catch (Exception)
                    {
                        error = true;
                        message =
                            "Please enter a valid value for terms of service code. The only options are 1 for full time and 2 for part time ";
                    }
                    String highestAcademicQualification = "";
                    try
                    {
                        highestAcademicQualification = String.IsNullOrEmpty(current[16].Trim())
                         ? ""
                         : current[16].Trim();
                    }
                    catch (Exception) {
                    }
                    int academicQualificationCode = 0;
                    try
                    {
                        academicQualificationCode = Convert.ToInt32(highestAcademicQualification);
                        if (academicQualificationCode > 7 || academicQualificationCode < 1)
                        {
                            throw new Exception();
                        }
                        academicQualificationCode -= 1;
                    }
                    catch (Exception)
                    {
                        error = true;
                        message =
                            "Please enter a valid value for highest academic qualification code. The options are between 1 and 7";
                    }
                    String campus = "";
                    try
                    {
                        campus = String.IsNullOrEmpty(current[17].Trim()) ? "" : current[17].Trim();
                    }
                    catch (Exception) { }
                    String userName = value.userUserName, password = value.userPassword;
                    if (error)
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {

                            errorRecord[i] = current[i];
                            //convert date of birth and first appointment date to human readable formats
                            if (i == 7 || i == 14)
                            {
                                try
                                {
                                    String myDate =
                                        DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                    // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                    errorRecord[i] = myDate;
                                }
                                catch (Exception)
                                {
                                    errorRecord[i] = current[i];
                                }
                            }
                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    else
                    {
                        String status = new Config().ObjNav()
                            .AddAcademicStaff(idNumber, payrollNo, firstName, middleName, lastName, genderCode,
                                ethnicBackground, dateOfBirth, nationalityCode, homeCounty,
                                disabilityDescription, disabilityRegistrationCode, domainCode, rankCode,
                                dateOfFirstAppointment, termsOfServiceCode, academicQualificationCode,campus,dateOfCurrentAppointment,
                                userName, password);

                        // added = false;
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                if (i == 7 || i == 14)
                                {
                                    try
                                    {
                                        String myDate =
                                            DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }

                }

                catch (Exception t)
                {
                    String errorMessage = t.Message;
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {

                        errorRecord[i] = current[i];
                        //convert date of birth and first appointment date to human readable formats
                        if (i == 7 || i == 14)
                        {
                            try
                            {
                                String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                errorRecord[i] = myDate;
                            }
                            catch (Exception)
                            {
                                errorRecord[i] = current[i];
                            }
                        }
                    }
                    errorRecord[current.Length] = errorMessage;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;
            return sj;//value; //value2 + "value2";
        }
        [Route("addPartTimeAcademicStaff")]
        public StudentJson AddPartTimeAcademicStaff([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;


            List<String[]> errorData = new List<string[]>();
            foreach (String[] current in data)
            {
                try
                {
                    String message = "";
                    Boolean error = false;
                    String idNumber = "";
                    try
                    {
                        idNumber = String.IsNullOrEmpty(current[0].Trim()) ? "" : current[0].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String payrollNo = "";
                    try
                    {
                        payrollNo = String.IsNullOrEmpty(current[1].Trim()) ? "" : current[1].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String firstName = "";
                    try
                    {
                        firstName = String.IsNullOrEmpty(current[2].Trim()) ? "" : current[2].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String middleName = "";
                    try
                    {
                        middleName = String.IsNullOrEmpty(current[3].Trim()) ? "" : current[3].Trim();
                    }
                    catch (Exception) { }
                    String lastName = "";
                    try
                    {
                        lastName = String.IsNullOrEmpty(current[4].Trim()) ? "" : current[4].Trim();
                    }
                    catch (Exception) { }
                    if (String.IsNullOrEmpty(idNumber))
                    {
                        error = true;
                        message = "ID No/Passport No is mandatory";
                    }
                    else if ((String.IsNullOrEmpty(firstName) && String.IsNullOrEmpty(middleName)) ||
                             (String.IsNullOrEmpty(middleName) && String.IsNullOrEmpty(lastName)) ||
                             (String.IsNullOrEmpty(lastName) && String.IsNullOrEmpty(firstName)))
                    {
                        error = true;
                        message = "A staff must have at least two names";
                    }
                    String gender = "";
                    try
                    {
                        gender = String.IsNullOrEmpty(current[5].Trim()) ? "" : current[5].Trim();
                    }
                    catch (Exception) { }
                    gender = gender.ToLower();
                    int genderCode = 1;
                    if (gender == "male")
                    {
                        genderCode = 0;
                    }
                    else if (gender == "female")
                    {
                        genderCode = 1;
                    }
                    else if (gender == "intersex")
                    {
                        genderCode = 2;
                    }
                    else
                    {
                        error = true;
                        message = "Please enter a valid gender. The only options are (male, female, intersex)";
                    }
                    String ethnicBackground = "";
                    try
                    {
                        ethnicBackground = String.IsNullOrEmpty(current[6].Trim()) ? "" : current[6].Trim();
                    }
                    catch (Exception) { }
                    String myDob = "";
                    try
                    {
                        myDob = String.IsNullOrEmpty(current[7].Trim()) ? "" : current[7].Trim();
                    }
                    catch (Exception) { }
                    DateTime dateOfBirth = new DateTime();
                    try
                    {
                        dateOfBirth = DateTime.FromOADate(Convert.ToDouble(myDob));
                        if (dateOfBirth > DateTime.Now)
                        {
                            error = true;
                            message = "Date of birth cannot be earlier than today";
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid value for date of birth";
                    }
                    String nationalityCode = "";
                    try
                    {
                        nationalityCode = String.IsNullOrEmpty(current[8].Trim()) ? "" : current[8].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String homeCounty = "";
                    try
                    {
                        homeCounty = String.IsNullOrEmpty(current[9].Trim()) ? "" : current[9].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String disabilityDescription = "";
                    try
                    {
                        disabilityDescription = String.IsNullOrEmpty(current[10].Trim()) ? "" : current[10].Trim();
                    }
                    catch (Exception) { }
                    String disabilityRegistrationCode = "";
                    try
                    {
                        disabilityRegistrationCode = String.IsNullOrEmpty(current[11].Trim())
                        ? ""
                        : current[11].Trim();
                    }
                    catch (Exception)
                    {
                    }
                    //String domainCode = "";
                    //try
                    //{
                    //    domainCode = String.IsNullOrEmpty(current[12].Trim()) ? "" : current[12].Trim();
                    //}
                    //catch (Exception)
                    //{
                    //}

                    String rankCode = "";
                    try
                    {
                        rankCode = String.IsNullOrEmpty(current[14].Trim()) ? "" : current[14].Trim();
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid value Staff rank";
                    }

                    //DateTime dateOfFirstAppointment = new DateTime();
                    //try
                    //{
                    //    dateOfFirstAppointment = DateTime.FromOADate(Convert.ToDouble(appointmentDate));
                    //    if (dateOfFirstAppointment > DateTime.Now)
                    //    {
                    //        error = true;
                    //        message = "Date of first appointment cannot be earlier than today";
                    //    }
                    //}
                    //catch (Exception)
                    //{
                    //    error = true;
                    //    message = "Please enter a valid value for date of first appointment";
                    //}
                    //String currentappointmentDate = "";
                    //try
                    //{
                    //    currentappointmentDate = String.IsNullOrEmpty(current[18].Trim()) ? "" : current[18].Trim();
                    //}
                    //catch (Exception)
                    //{
                    //}
                    //DateTime dateOfCurrentAppointment = new DateTime();
                    //try
                    //{
                    //    dateOfCurrentAppointment = DateTime.FromOADate(Convert.ToDouble(currentappointmentDate));
                    //    if (dateOfCurrentAppointment > DateTime.Now)
                    //    {
                    //        error = true;
                    //        message = "Date of first Current appointment cannot be earlier than today";
                    //    }
                    //}
                    //catch (Exception)
                    //{
                    //    error = true;
                    //    message = "Please enter a valid value for Date of Current appointment";
                    //}
                    String stafftype = "";
                    try
                    {
                        stafftype = String.IsNullOrEmpty(current[15].Trim()) ? "" : current[15].Trim();
                    }
                    catch (Exception)
                    {
                    }
                    int category = 0;
                    try
                    {
                        category = Convert.ToInt32(stafftype);
                        if (category > 2 || category < 1)
                        {
                            throw new Exception();
                        }
                        category -= 1;
                    }
                    catch (Exception)
                    {
                        error = true;
                        message =
                            "Please enter a valid value for staff category code. The only options are 2 for Internal and 1 for External ";
                    }
                    String highestAcademicQualification = "";
                    try
                    {
                        highestAcademicQualification = String.IsNullOrEmpty(current[13].Trim()) ? "" : current[13].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    int academicQualificationCode = 0;
                    try
                    {
                        academicQualificationCode = Convert.ToInt32(highestAcademicQualification);
                        if (academicQualificationCode > 7 || academicQualificationCode < 1)
                        {
                            throw new Exception();
                        }
                        academicQualificationCode -= 1;
                    }
                    catch (Exception)
                    {
                        error = true;
                        message =
                            "Please enter a valid value for highest academic qualification code. The options are between 1 and 7";
                    }
                    String campus = "";
                    try
                    {
                        campus = String.IsNullOrEmpty(current[12].Trim()) ? "" : current[12].Trim();
                    }
                    catch (Exception) { }
                    String userName = value.userUserName, password = value.userPassword;
                    if (error)
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {

                            errorRecord[i] = current[i];
                            //convert date of birth and first appointment date to human readable formats
                            if (i == 7 || i == 14)
                            {
                                try
                                {
                                    String myDate =
                                        DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                    // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                    errorRecord[i] = myDate;
                                }
                                catch (Exception)
                                {
                                    errorRecord[i] = current[i];
                                }
                            }
                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    else
                    {
                        String status = new Config().ObjNav()
                            .AddPartTimeAcademicStaff(idNumber, payrollNo, firstName, middleName, lastName, genderCode,
                                ethnicBackground, dateOfBirth, nationalityCode, homeCounty,
                                disabilityDescription, disabilityRegistrationCode, campus, academicQualificationCode, rankCode, category, userName, password);

                        // added = false;
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                if (i == 7 || i == 14)
                                {
                                    try
                                    {
                                        String myDate =
                                            DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }

                }

                catch (Exception t)
                {
                    String errorMessage = t.Message;
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {

                        errorRecord[i] = current[i];
                        //convert date of birth and first appointment date to human readable formats
                        if (i == 7 || i == 14)
                        {
                            try
                            {
                                String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                errorRecord[i] = myDate;
                            }
                            catch (Exception)
                            {
                                errorRecord[i] = current[i];
                            }
                        }
                    }
                    errorRecord[current.Length] = errorMessage;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;
            return sj;//value; //value2 + "value2";
        }
        [Route("addLibraryStaff")]
        public StudentJson AddLibraryStaff([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;


            List<String[]> errorData = new List<string[]>();
            foreach (String[] current in data)
            {
                try
                {
                    String message = "";
                    Boolean error = false;
                    String idNumber = "";
                    String firstName = "";
                    String middleName = "";
                    String lastName = "";
                    String position = "";
                    String highestAcademicQualification = "";
                    String campusCode = "";
                    try { idNumber = String.IsNullOrEmpty(current[0].Trim()) ? "" : current[0].Trim(); } catch (Exception) { }
                    try { firstName = String.IsNullOrEmpty(current[1].Trim()) ? "" : current[1].Trim(); } catch (Exception) { }
                    try { middleName = String.IsNullOrEmpty(current[2].Trim()) ? "" : current[2].Trim(); } catch (Exception) { }
                    try { lastName = String.IsNullOrEmpty(current[3].Trim()) ? "" : current[3].Trim(); } catch (Exception) { }

                    try { position = String.IsNullOrEmpty(current[5].Trim()) ? "" : current[5].Trim(); } catch (Exception) { }
                    try { highestAcademicQualification = String.IsNullOrEmpty(current[6].Trim()) ? "" : current[6].Trim(); } catch (Exception) { }
                    try { campusCode = String.IsNullOrEmpty(current[7].Trim()) ? "" : current[7].Trim(); } catch (Exception) { }




                    if (String.IsNullOrEmpty(idNumber))
                    {
                        error = true;
                        message = "ID No/Passport No is mandatory";
                    }
                    else if ((String.IsNullOrEmpty(firstName) && String.IsNullOrEmpty(middleName)) ||
                             (String.IsNullOrEmpty(middleName) && String.IsNullOrEmpty(lastName)) ||
                             (String.IsNullOrEmpty(lastName) && String.IsNullOrEmpty(firstName)))
                    {
                        error = true;
                        message = "A staff must have at least two names";
                    }


                    String myDob = "";
                    try
                    {
                        myDob = String.IsNullOrEmpty(current[4].Trim()) ? "" : current[4].Trim();
                    }
                    catch (Exception) { }
                    DateTime dateOfBirth = new DateTime();
                    try
                    {
                        dateOfBirth = DateTime.FromOADate(Convert.ToDouble(myDob));
                        if (dateOfBirth > DateTime.Now)
                        {
                            error = true;
                            message = "Date of birth cannot be earlier than today";
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid value for date of birth";
                    }

                    int academicQualificationCode = 0;
                    try
                    {
                        academicQualificationCode = Convert.ToInt32(highestAcademicQualification);
                        if (academicQualificationCode > 7 || academicQualificationCode < 1)
                        {
                            throw new Exception();
                        }
                        academicQualificationCode -= 1;
                    }
                    catch (Exception)
                    {
                        error = true;
                        message =
                            "Please enter a valid value for highest academic qualification code. The options are between 1 and 7";
                    }

                    String userName = value.userUserName, password = value.userPassword;
                    if (error)
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {

                            errorRecord[i] = current[i];
                            //convert date of birth 
                            if (i == 4)
                            {
                                try
                                {
                                    String myDate =
                                        DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                    // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                    errorRecord[i] = myDate;
                                }
                                catch (Exception)
                                {
                                    errorRecord[i] = current[i];
                                }
                            }
                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    else
                    {
                        String status = new Config().ObjNav()
                            .AddLibraryStaff(userName, password, idNumber, firstName, middleName, lastName, dateOfBirth, position, academicQualificationCode, campusCode);

                        // added = false;
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                if (i == 7 || i == 14)
                                {
                                    try
                                    {
                                        String myDate =
                                            DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }

                }

                catch (Exception t)
                {
                    String errorMessage = t.Message;
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {

                        errorRecord[i] = current[i];
                        //convert date of birth and first appointment date to human readable formats
                        if (i == 4)
                        {
                            try
                            {
                                String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                errorRecord[i] = myDate;
                            }
                            catch (Exception)
                            {
                                errorRecord[i] = current[i];
                            }
                        }
                    }
                    errorRecord[current.Length] = errorMessage;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;
            return sj;//value; //value2 + "value2";
        }

        [Route("addNonAcademicStaff")]
        public StudentJson AddNonAcademicStaff([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;


            List<String[]> errorData = new List<string[]>();
            foreach (String[] current in data)
            {
                try
                {
                    String message = "";
                    Boolean error = false;
                    String idNumber = "";
                    try
                    {
                        idNumber = String.IsNullOrEmpty(current[0].Trim()) ? "" : current[0].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String payrollNo = "";
                    try
                    {
                        payrollNo = String.IsNullOrEmpty(current[1].Trim()) ? "" : current[1].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String firstName = "";
                    try
                    {
                        firstName = String.IsNullOrEmpty(current[2].Trim()) ? "" : current[2].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String middleName = "";
                    try
                    {
                        middleName = String.IsNullOrEmpty(current[3].Trim()) ? "" : current[3].Trim();
                    }
                    catch (Exception) { }
                    String lastName = "";
                    try
                    {
                        lastName = String.IsNullOrEmpty(current[4].Trim()) ? "" : current[4].Trim();
                    }
                    catch (Exception) {

                    }
                    if (String.IsNullOrEmpty(idNumber))
                    {
                        error = true;
                        message = "ID No/Passport No is mandatory";
                    }
                    else if ((String.IsNullOrEmpty(firstName) && String.IsNullOrEmpty(middleName)) ||
                             (String.IsNullOrEmpty(middleName) && String.IsNullOrEmpty(lastName)) ||
                             (String.IsNullOrEmpty(lastName) && String.IsNullOrEmpty(firstName)))
                    {
                        error = true;
                        message = "A staff must have at least two names";
                    }
                    String gender = "";
                    try
                    {
                        gender = String.IsNullOrEmpty(current[5].Trim()) ? "" : current[5].Trim();
                    }
                    catch (Exception) {

                    }
                    gender = gender.ToLower();
                    int genderCode = 1;
                    if (gender == "male")
                    {
                        genderCode = 0;
                    }
                    else if (gender == "female")
                    {
                        genderCode = 1;
                    }
                    else if (gender == "intersex")
                    {
                        genderCode = 2;
                    }
                    else
                    {
                        error = true;
                        message = "Please enter a valid gender. The only options are (male, female, intersex)";
                    }
                    String ethnicBackground = "";
                    try
                    {
                        ethnicBackground = String.IsNullOrEmpty(current[6].Trim()) ? "" : current[6].Trim();
                    }
                    catch (Exception) { }
                    String myDob = "";
                    try
                    {
                        myDob = String.IsNullOrEmpty(current[7].Trim()) ? "" : current[7].Trim();
                    }
                    catch (Exception) {

                    }
                    DateTime dateOfBirth = new DateTime();
                    try
                    {
                        dateOfBirth = DateTime.FromOADate(Convert.ToDouble(myDob));
                        if (dateOfBirth > DateTime.Now)
                        {
                            error = true;
                            message = "Date of birth cannot be earlier than today";
                        }
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid value for date of birth";
                    }
                    String nationalityCode = "";
                    try
                    {
                        nationalityCode = String.IsNullOrEmpty(current[8].Trim()) ? "" : current[8].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String homeCounty = "";
                    try
                    {
                        homeCounty = String.IsNullOrEmpty(current[9].Trim()) ? "" : current[9].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String disabilityDescription = "";
                    try
                    {
                        disabilityDescription = String.IsNullOrEmpty(current[10].Trim()) ? "" : current[10].Trim();
                    }
                    catch (Exception) {  
                    }
                    String disabilityRegistrationCode = "";
                    try
                    {
                        disabilityRegistrationCode = String.IsNullOrEmpty(current[11].Trim())
                        ? ""
                        : current[11].Trim();
                    }
                    catch (Exception)
                    {
                    }
                    String campus = "";
                    try
                    {
                        campus = String.IsNullOrEmpty(current[12].Trim()) ? "" : current[12].Trim();
                    }
                    catch (Exception)
                    {

                    }
                    String highestAcademicQualification = "";
                    try
                    {
                        highestAcademicQualification = String.IsNullOrEmpty(current[13].Trim())? "" : current[13].Trim();
                    }
                    catch (Exception)
                    {
                    }
                    int academicQualificationCode = 0;
                    try
                    {
                        academicQualificationCode = Convert.ToInt32(highestAcademicQualification);
                        if (academicQualificationCode > 7 || academicQualificationCode < 1)
                        {
                            throw new Exception();
                        }
                        academicQualificationCode -= 1;
                    }
                    catch (Exception)
                    {
                        error = true;
                        message =
                            "Please enter a valid value for highest academic qualification code. The options are between 1 and 7";
                    }

                    //String domainCode = "";
                    //try
                    //{
                    //    domainCode = String.IsNullOrEmpty(current[12].Trim()) ? "" : current[12].Trim();
                    //}
                    //catch (Exception)
                    //{
                    //}

                    String rankCode = "";
                    try
                    {
                        rankCode = String.IsNullOrEmpty(current[14].Trim()) ? "" : current[14].Trim();
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please enter a valid value Staff rank";
                    }
                    //String appointmentDate = "";
                    //try
                    //{
                    //    appointmentDate = String.IsNullOrEmpty(current[14].Trim()) ? "" : current[14].Trim();
                    //}
                    //catch (Exception)
                    //{
                    //}
                    //DateTime dateOfFirstAppointment = new DateTime();
                    //try
                    //{
                    //    dateOfFirstAppointment = DateTime.FromOADate(Convert.ToDouble(appointmentDate));
                    //    if (dateOfFirstAppointment > DateTime.Now)
                    //    {
                    //        error = true;
                    //        message = "Date of first appointment cannot be earlier than today";
                    //    }
                    //}
                    //catch (Exception)
                    //{
                    //    error = true;
                    //    message = "Please enter a valid value for date of first appointment";
                    //}
                    //String currentappointmentDate = "";
                    //try
                    //{
                    //    currentappointmentDate = String.IsNullOrEmpty(current[18].Trim()) ? "" : current[18].Trim();
                    //}
                    //catch (Exception)
                    //{
                    //}
                    //DateTime dateOfCurrentAppointment = new DateTime();
                    //try
                    //{
                    //    dateOfCurrentAppointment = DateTime.FromOADate(Convert.ToDouble(currentappointmentDate));
                    //    if (dateOfCurrentAppointment > DateTime.Now)
                    //    {
                    //        error = true;
                    //        message = "Date of  Current appointment cannot be earlier than today";
                    //    }
                    //}
                    //catch (Exception)
                    //{
                    //    error = true;
                    //    message = "Please enter a valid value for Date of Current appointment";
                    //}
                    //String termsOfService = "";
                    //try
                    //{
                    //    termsOfService = String.IsNullOrEmpty(current[15].Trim()) ? "" : current[15].Trim();
                    //}
                    //catch (Exception)
                    //{
                    //}
                    //int termsOfServiceCode = 0;
                    //try
                    //{
                    //    termsOfServiceCode = Convert.ToInt32(termsOfService);
                    //    if (termsOfServiceCode > 2 || termsOfServiceCode < 1)
                    //    {
                    //        throw new Exception();
                    //    }
                    //    termsOfServiceCode -= 1;
                    //}
                    //catch (Exception)
                    //{
                    //    error = true;
                    //    message =
                    //        "Please enter a valid value for terms of service code. The only options are 1 for full time and 2 for part time ";
                    //}
                  
                    String userName = value.userUserName, password = value.userPassword;
                    if (error)
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {

                            errorRecord[i] = current[i];
                            //convert date of birth and first appointment date to human readable formats
                            if (i == 7 || i == 14)
                            {
                                try
                                {
                                    String myDate =
                                        DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                    // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                    errorRecord[i] = myDate;
                                }
                                catch (Exception)
                                {
                                    errorRecord[i] = current[i];
                                }
                            }
                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    else
                    {
                        String status = new Config().ObjNav()
                            .AddNonAcademicStaff(idNumber, payrollNo, firstName, middleName, lastName, genderCode,ethnicBackground, dateOfBirth, nationalityCode, homeCounty,
                                disabilityDescription, disabilityRegistrationCode, campus, academicQualificationCode, rankCode,userName, password);

                        // added = false;
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                if (i == 7 || i == 14)
                                {
                                    try
                                    {
                                        String myDate =
                                            DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }

                }

                catch (Exception t)
                {
                    String errorMessage = t.Message;
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {

                        errorRecord[i] = current[i];
                        //convert date of birth and first appointment date to human readable formats
                        if (i == 7 || i == 14)
                        {
                            try
                            {
                                String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                errorRecord[i] = myDate;
                            }
                            catch (Exception)
                            {
                                errorRecord[i] = current[i];
                            }
                        }
                    }
                    errorRecord[current.Length] = errorMessage;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.data = errorData;
            sj.column_map = value.column_map;
            return sj;//value; //value2 + "value2";
        }
        [Route("addCase")]
        public StudentJson AddCase([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            Boolean added = true;

            List<String[]> errorData = new List<string[]>();
            foreach (String[] current in data)
            {
                try
                {
                    /*
                 Case Description	Case Date	Verdict	Verdict Date

                    */
                    String message = "";
                    Boolean error = false;
                    String studentId = "";
                    try
                    {
                        studentId = current[0].Trim();
                    }
                    catch (Exception)
                    {
                        studentId = "";
                    }
                    String admissionNo = "";
                    try
                    {
                        admissionNo = current[1].Trim();
                    }
                    catch (Exception)
                    {
                        admissionNo = "";
                    }
                    String firstName = "";
                    try
                    {
                        firstName = current[2].Trim();
                    }
                    catch (Exception)
                    {
                        firstName = "";
                    }
                    String middleName = "";
                    try
                    {
                        middleName = current[3].Trim();
                    }
                    catch (Exception)
                    {
                        middleName = "";
                    }
                    String lastName = "";
                    try
                    {
                        lastName = current[4].Trim();
                    }
                    catch (Exception)
                    {
                        lastName = "";
                    }
                    String caseDescription = "";
                    try
                    {
                        caseDescription = current[5].Trim();
                    }
                    catch (Exception)
                    {
                        caseDescription = "";
                    }
                    DateTime caseDate = new DateTime();
                    try
                    {
                        caseDate = DateTime.FromOADate(Convert.ToDouble(current[6].Trim()));
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please give a valid case date";
                    }
                    String verdict = "";
                    try
                    {
                        verdict = current[7].Trim();
                    }
                    catch (Exception)
                    {
                        verdict = "";
                    }
                    DateTime verdictDate = new DateTime();
                    try
                    {
                        verdictDate = DateTime.FromOADate(Convert.ToDouble(current[8]));
                    }
                    catch (Exception)
                    {
                        error = true;
                        message += "Please give a valid verdict date";
                    }
                    String userName = value.userUserName, password = value.userPassword;
                    if (error)
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {

                            errorRecord[i] = current[i];
                            //convert date of birth and first appointment date to human readable formats
                            if (i == 6 || i == 8)
                            {
                                try
                                {
                                    String myDate =
                                        DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                    // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                    errorRecord[i] = myDate;
                                }
                                catch (Exception)
                                {
                                    errorRecord[i] = current[i];
                                }
                            }
                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    else
                    {
                        String status = new Config().ObjNav()
                            .AddDisciplineCase(studentId, admissionNo, firstName, middleName, lastName, caseDescription, caseDate, verdict, verdictDate, userName,
                                password);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                if (i == 6 || i == 8)
                                {
                                    try
                                    {
                                        String myDate =
                                            DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    //value2 = value.column_map;
                }
                catch (Exception t)
                {
                    String errorMessage = t.Message;
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {

                        errorRecord[i] = current[i];
                        //convert date of birth and first appointment date to human readable formats
                        if (i == 6 || i == 8)
                        {
                            try
                            {
                                String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                errorRecord[i] = myDate;
                            }
                            catch (Exception)
                            {
                                errorRecord[i] = current[i];
                            }
                        }
                    }
                    errorRecord[current.Length] = errorMessage;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.column_map = "";
            sj.data = errorData;

            return sj; //value2 + "value2";
        }
        [Route("addAppealsCase")]
        public StudentJson AddAppealsCase([FromBody]StudentJson value)
        {
            List<String[]> data = value.data;
            Boolean added = true;

            List<String[]> errorData = new List<string[]>();
            foreach (String[] current in data)
            {
                try
                {
                    /*
                 Case Description	Case Date	Verdict	Verdict Date

                    */
                    String message = "";
                    Boolean error = false;
                    String studentId = "";
                    try
                    {
                        studentId = current[0].Trim();
                    }
                    catch (Exception)
                    {
                        studentId = "";
                    }
                    String admissionNo = "";
                    try
                    {
                        admissionNo = current[1].Trim();
                    }
                    catch (Exception)
                    {
                        admissionNo = "";
                    }
                    String firstName = "";
                    try
                    {
                        firstName = current[2].Trim();
                    }
                    catch (Exception)
                    {
                        firstName = "";
                    }
                    String middleName = "";
                    try
                    {
                        middleName = current[3].Trim();
                    }
                    catch (Exception)
                    {
                        middleName = "";
                    }
                    String lastName = "";
                    try
                    {
                        lastName = current[4].Trim();
                    }
                    catch (Exception)
                    {
                        lastName = "";
                    }
                    String caseDescription = "";
                    try
                    {
                        caseDescription = current[5].Trim();
                    }
                    catch (Exception)
                    {
                        caseDescription = "";
                    }
                    DateTime caseDate = new DateTime();
                    try
                    {
                        caseDate = DateTime.FromOADate(Convert.ToDouble(current[6].Trim()));
                    }
                    catch (Exception)
                    {
                        error = true;
                        message = "Please give a valid case date";
                    }
                    String verdict = "";
                    try
                    {
                        verdict = current[7].Trim();
                    }
                    catch (Exception)
                    {
                        verdict = "";
                    }
                    DateTime verdictDate = new DateTime();
                    try
                    {
                        verdictDate = DateTime.FromOADate(Convert.ToDouble(current[8]));
                    }
                    catch (Exception)
                    {
                        error = true;
                        message += "Please give a valid verdict date";
                    }
                    String userName = value.userUserName, password = value.userPassword;
                    if (error)
                    {
                        String[] errorRecord = new string[current.Length + 1];
                        for (int i = 0; i < current.Length; i++)
                        {

                            errorRecord[i] = current[i];
                            //convert date of birth and first appointment date to human readable formats
                            if (i == 6 || i == 8)
                            {
                                try
                                {
                                    String myDate =
                                        DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                    // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                    errorRecord[i] = myDate;
                                }
                                catch (Exception)
                                {
                                    errorRecord[i] = current[i];
                                }
                            }
                        }
                        errorRecord[current.Length] = message;
                        errorData.Add(errorRecord);
                    }
                    else
                    {
                        String status = new Config().ObjNav()
                            .AddSuccessfulAppeals(studentId, admissionNo, firstName, middleName, lastName, caseDescription, caseDate, verdict, verdictDate, userName,
                                password);
                        if (status != "success")
                        {
                            String[] errorRecord = new string[current.Length + 1];
                            for (int i = 0; i < current.Length; i++)
                            {

                                errorRecord[i] = current[i];
                                if (i == 6 || i == 8)
                                {
                                    try
                                    {
                                        String myDate =
                                            DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                        errorRecord[i] = myDate;
                                    }
                                    catch (Exception)
                                    {
                                        errorRecord[i] = current[i];
                                    }
                                }
                            }
                            errorRecord[current.Length] = status;
                            errorData.Add(errorRecord);

                        }
                    }
                    //value2 = value.column_map;
                }
                catch (Exception t)
                {
                    String errorMessage = t.Message;
                    String[] errorRecord = new string[current.Length + 1];
                    for (int i = 0; i < current.Length; i++)
                    {

                        errorRecord[i] = current[i];
                        //convert date of birth and first appointment date to human readable formats
                        if (i == 6 || i == 8)
                        {
                            try
                            {
                                String myDate = DateTime.FromOADate(Convert.ToDouble(current[i])).ToString("dd-MM-yy");
                                // myDate = myDate.Replace("-", "/").Trim(); //'2018-07-22' 29/08/2019
                                errorRecord[i] = myDate;
                            }
                            catch (Exception)
                            {
                                errorRecord[i] = current[i];
                            }
                        }
                    }
                    errorRecord[current.Length] = errorMessage;
                    errorData.Add(errorRecord);
                }
            }

            StudentJson sj = new StudentJson();
            sj.column_map = "";
            sj.data = errorData;

            return sj; //value2 + "value2";
        }
        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }
        // http://41.89.47.15:7052/CUENAV/WS/CUE/Codeunit/CuePortal
        // DELETE api/values/5
        public void Delete(int id)
        {
        }
    }
}
