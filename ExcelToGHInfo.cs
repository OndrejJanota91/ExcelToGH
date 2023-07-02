using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace ExcelToGH
{
    public class ExcelToGHInfo : GH_AssemblyInfo
    {
        public override string Name => "ExcelToGH";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "Simple .gha file with one component which reads excel file";

        public override Guid Id => new Guid("37AF2F0B-58CA-425F-9E4C-77B02CC751ED");

        //Return a string identifying you or your company.
        public override string AuthorName => "Ondřej Janota";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "janotaondrej91@gmail.com";
    }
}