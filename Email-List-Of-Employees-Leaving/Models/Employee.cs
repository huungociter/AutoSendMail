using System;

namespace EmployeeSendMailModels
{
    public class Employee
    {
        public string total { get; set; }
        public List<Result> result { get; set; }

        public class Result
        {
            public string firstName { get; set; }
            public string lastName { get; set; }
            public string accountName { get; set; }
            public string email { get; set; }
            public string employeeCode { get; set; }
            public string unitName { get; set; }
            public string unitCode { get; set; }
            public string divisionName { get; set; }
            public string divisionCode { get; set; }
            public string departmentName { get; set; }
            public string departmentCode { get; set; }
            public string sectionName { get; set; }
            public string sectionCode { get; set; }
            public string groupName { get; set; }
            public string groupCode { get; set; }
            public string costCenterName { get; set; }
            public string costCenter { get; set; }
            public string positionName { get; set; }
            public string positionCode { get; set; }
            public DateTime? ngayVaoLam { get; set; }
            public DateTime? ngayKetThucThuViec { get; set; }
            public DateTime? ngayLamChinhThuc { get; set; }
            public DateTime? ngayNghiViec { get; set; }
            public bool active { get; set; }
            public string maTinhTrang { get; set; }
            public string tenTinhTrang { get; set; }
            public bool isCa3 { get; set; }
        }
    }
}