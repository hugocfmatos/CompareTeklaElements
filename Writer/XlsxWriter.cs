using Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Writer
{
    public class XlsxWriter
    {
        public static void Generate(List<Member> members, string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new();

            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Members");

            BuildHeader(worksheet);

            FillBody(worksheet, members);

            File.WriteAllBytes(path, package.GetAsByteArray());
        }

        private static void BuildHeader(ExcelWorksheet worksheet)
        {
            worksheet.Cells[2, 2].Value = "GUID - First Model";
            worksheet.Cells[2, 3].Value = "Start Member Coordinates - First Model";
            worksheet.Cells[2, 4].Value = "End Member Coordinates - First Model";
            worksheet.Cells[2, 5].Value = "Profile - First Model";
            worksheet.Cells[2, 6].Value = "Assembly Position - First Model";
            worksheet.Cells[2, 7].Value = "GUID - Second Model";
            worksheet.Cells[2, 8].Value = "Start Member Coordinates - Second Model";
            worksheet.Cells[2, 9].Value = "End Member Coordinates - Second Model";
            worksheet.Cells[2, 10].Value = "Profile - Second Model";
            worksheet.Cells[2, 11].Value = "Assembly Position - Second Model";
            worksheet.Cells[2, 12].Value = "Check Result";
            worksheet.Cells[2, 13].Value = "Remarks";
        }

        private static void FillBody(ExcelWorksheet worksheet, List<Member> members)
        {
            for (int i = 1; i < members.Count; i++)
            {
                worksheet.Cells[i + 2, 2].Value = members[i].GUID1;
                worksheet.Cells[i + 2, 3].Value = UniteStartCoordinates1(members[i]);
                worksheet.Cells[i + 2, 4].Value = UniteEndCoordinates1(members[i]);
                worksheet.Cells[i + 2, 5].Value = members[i].Profile1;
                worksheet.Cells[i + 2, 6].Value = members[i].AssemblyPosition1;
                worksheet.Cells[i + 2, 7].Value = members[i].GUID2;
                worksheet.Cells[i + 2, 8].Value = UniteStartCoordinates2(members[i]);
                worksheet.Cells[i + 2, 9].Value = UniteEndCoordinates2(members[i]);
                worksheet.Cells[i + 2, 10].Value = members[i].Profile2;
                worksheet.Cells[i + 2, 11].Value = members[i].AssemblyPosition2;
                worksheet.Cells[i + 2, 12].Value = members[i].Result;
                worksheet.Cells[i + 2, 13].Value = members[i].Remarks;
                ColourRow(members[i], worksheet, i);
            }
        }

        // Instead of this 4 methods, 1 could be used using reflection. This option was not adopted due to worse performance
        private static string UniteStartCoordinates1(Member member)
        {
            if (member.StartMemberCoordinatesX1 == 0)
                return "";
            else
            return $"\"({member.StartMemberCoordinatesX1}, {member.StartMemberCoordinatesY1}, {member.StartMemberCoordinatesZ1})\"";
        }

        private static string UniteEndCoordinates1(Member member)
        {
            if (member.EndMemberCoordinatesX1 == 0)
                return "";
            else
                return $"\"({member.EndMemberCoordinatesX1}, {member.EndMemberCoordinatesY1}, {member.EndMemberCoordinatesZ1})\"";
        }

        private static string UniteStartCoordinates2(Member member)
        {
            if (member.StartMemberCoordinatesX2 == 0)
                return "";
            else
                return $"\"({member.StartMemberCoordinatesX2}, {member.StartMemberCoordinatesY2}, {member.StartMemberCoordinatesZ2})\"";
        }

        private static string UniteEndCoordinates2(Member member)
        {
            if (member.EndMemberCoordinatesX2 == 0)
                return "";
            else
                return $"\"({member.EndMemberCoordinatesX2}, {member.EndMemberCoordinatesY2}, {member.EndMemberCoordinatesZ2})\"";
        }

        private static void ColourRow(Member member, ExcelWorksheet worksheet, int index)
        {
            Color color = member.Result switch
            {
                "New" => Color.Blue,
                "Deleted" => Color.Red,
                "Modified" => Color.Yellow,
                _ => Color.Green,
            };
            worksheet.Cells[$"B{index + 2}:M{index + 2}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[$"B{index + 2}:M{index + 2}"].Style.Fill.BackgroundColor.SetColor(color);
        }
    }
}