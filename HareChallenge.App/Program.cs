using Core.Models;
using Reader;
using System.Data;
using Writer;

class Program
{
    static void Main(string[] args)
    {
        const string model1Path = "C:\\Model1 data.csv";
        const string model2Path = "C:\\Model2 data.csv";

        DataTable memberList1 = CsvReader.Read(model1Path);

        DataTable memberList2 = CsvReader.Read(model2Path);

        List<Member> members = SortMembers(memberList1, memberList2);

        XlsxWriter.Generate(members, "C:\\MemberList.xlsx");
    }

    public static List<Member> SortMembers(DataTable memberList1, DataTable memberList2)
    {
        List<Member> members = new();

        foreach(DataRow row in memberList1.Rows)
        {
            members.Add(BuildMember1(row));
        }

        foreach(DataRow row2 in memberList2.Rows)
        {
            foreach(Member member1 in members)
            {
                if (HasSameCoordinates(member1, row2))
                {
                    CompleteMember(member1, row2);
                    goto OuterLoop;
                }
            }

        members.Add(BuildMember2(row2));
        OuterLoop:
            continue;
        }

        foreach(Member member in members)
        {
            if (string.IsNullOrEmpty(member.Result))
                member.Result = "Deleted";
        }

        return members;
    }

    private static Member BuildMember1(DataRow row)
    {
        return new Member()
        {
            GUID1 = row["GUID"].ToString(),
            StartMemberCoordinatesX1 = (float)row["StartMemberCoordinatesX"],
            StartMemberCoordinatesY1 = (float)row["StartMemberCoordinatesY"],
            StartMemberCoordinatesZ1 = (float)row["StartMemberCoordinatesZ"],
            EndMemberCoordinatesX1 = (float)row["EndMemberCoordinatesX"],
            EndMemberCoordinatesY1 = (float)row["EndMemberCoordinatesY"],
            EndMemberCoordinatesZ1 = (float)row["EndMemberCoordinatesZ"],
            Profile1 = row["Profile"].ToString(),
            AssemblyPosition1 = (int)row["AssemblyPosition"]
        };
    }

    private static Member BuildMember2(DataRow row)
    {
        return new Member()
        {
            GUID2 = row["GUID"].ToString(),
            StartMemberCoordinatesX2 = (float)row["StartMemberCoordinatesX"],
            StartMemberCoordinatesY2 = (float)row["StartMemberCoordinatesY"],
            StartMemberCoordinatesZ2 = (float)row["StartMemberCoordinatesZ"],
            EndMemberCoordinatesX2 = (float)row["EndMemberCoordinatesX"],
            EndMemberCoordinatesY2 = (float)row["EndMemberCoordinatesY"],
            EndMemberCoordinatesZ2 = (float)row["EndMemberCoordinatesZ"],
            Profile2 = row["Profile"].ToString(),
            AssemblyPosition2 = (int)row["AssemblyPosition"],
            Result = "New"
        };
    }

    private static Boolean HasSameCoordinates(Member member1, DataRow member2)
    {
        return member1.StartMemberCoordinatesX1 >= (float)member2["StartMemberCoordinatesX"] - 50
            && member1.StartMemberCoordinatesX1 <= (float)member2["StartMemberCoordinatesX"] + 50
            && member1.StartMemberCoordinatesY1 >= (float)member2["StartMemberCoordinatesY"] - 50
            && member1.StartMemberCoordinatesY1 <= (float)member2["StartMemberCoordinatesY"] + 50
            && member1.StartMemberCoordinatesZ1 >= (float)member2["StartMemberCoordinatesZ"] - 50
            && member1.StartMemberCoordinatesZ1 <= (float)member2["StartMemberCoordinatesZ"] + 50
            && member1.EndMemberCoordinatesX1 >= (float)member2["EndMemberCoordinatesX"] - 50
            && member1.EndMemberCoordinatesX1 <= (float)member2["EndMemberCoordinatesX"] + 50
            && member1.EndMemberCoordinatesY1 >= (float)member2["EndMemberCoordinatesY"] - 50
            && member1.EndMemberCoordinatesY1 <= (float)member2["EndMemberCoordinatesY"] + 50
            && member1.EndMemberCoordinatesZ1 >= (float)member2["EndMemberCoordinatesZ"] - 50
            && member1.EndMemberCoordinatesZ1 <= (float)member2["EndMemberCoordinatesZ"] + 50;
    }

    private static void CompleteMember(Member member, DataRow row)
    {
        member.GUID2 = row["GUID"].ToString();
        member.StartMemberCoordinatesX2 = (float)row["StartMemberCoordinatesX"];
        member.StartMemberCoordinatesY2 = (float)row["StartMemberCoordinatesY"];
        member.StartMemberCoordinatesZ2 = (float)row["StartMemberCoordinatesZ"];
        member.EndMemberCoordinatesX2 = (float)row["EndMemberCoordinatesX"];
        member.EndMemberCoordinatesY2 = (float)row["EndMemberCoordinatesY"];
        member.EndMemberCoordinatesZ2 = (float)row["EndMemberCoordinatesZ"];
        member.Profile2 = row["Profile"].ToString();
        member.AssemblyPosition2 = (int)row["AssemblyPosition"];
        ComparePositionAndProfile(member);
    }

    private static void ComparePositionAndProfile(Member member)
    {
        if (member.Profile1.Equals(member.Profile2) && member.AssemblyPosition1 == member.AssemblyPosition2)
            member.Result = "Checked - OK";
        else
        {
            member.Result = "Modified";
            if (member.Profile1.Equals(member.Profile2) && member.AssemblyPosition1 != member.AssemblyPosition2)
                member.Remarks = "Assembly Position Changed";
            else
            {
                if (member.AssemblyPosition1 == member.AssemblyPosition2)
                    member.Remarks = "Profile Changed";
                else
                    member.Remarks = "Both Changed";
            }
        }
    }
}
