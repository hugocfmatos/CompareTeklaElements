namespace Core.Models
{
    public class Member
    {
        public string GUID1 { get; set; }
        public string GUID2 { get; set; }
        public float StartMemberCoordinatesX1 { get; set; }
        public float StartMemberCoordinatesY1 { get; set; }
        public float StartMemberCoordinatesZ1 { get; set; }
        public float StartMemberCoordinatesX2 { get; set; }
        public float StartMemberCoordinatesY2 { get; set; }
        public float StartMemberCoordinatesZ2 { get; set; }
        public float EndMemberCoordinatesX1 { get; set; }
        public float EndMemberCoordinatesY1 { get; set; }
        public float EndMemberCoordinatesZ1 { get; set; }
        public float EndMemberCoordinatesX2 { get; set; }
        public float EndMemberCoordinatesY2 { get; set; }
        public float EndMemberCoordinatesZ2 { get; set; }
        public string Profile1 { get; set; }
        public string Profile2 { get; set; }
        public int AssemblyPosition1 { get; set; }
        public int AssemblyPosition2 { get; set; }
        public string Result { get; set; }
        public string Remarks { get; set; }
    }
}