using RootLogic.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RootLogic
{
    public class LogicService
    {
        private List<string> FirstFileColumn { get; set; } = new List<string>(); //source
        private List<string> SecondFileColumn { get; set; } = new List<string>(); //logs

        public LogicService(string firstPath, int firstColumnNumber, string secondPath, int secondColumnNumber)
        {
            var firstStream = FileHelper.BuildStream(firstPath);
            var firstPosition = Tuple.Create(2, firstColumnNumber);
            FirstFileColumn = firstStream.GetColumnDataFromExcelFile(0, firstPosition);
            firstStream.Dispose();

            var secondStream = FileHelper.BuildStream(secondPath);
            var secondPosition = Tuple.Create(2, secondColumnNumber);
            SecondFileColumn = secondStream.GetColumnDataFromExcelFile(0, secondPosition);

        }

        private List<string> DiffLists()
        {
            return FirstFileColumn.Except(SecondFileColumn).ToList();
        }

        public void DiffAndSave(string path)
        {
            var diffs = DiffLists();
            diffs.WriteToFile(path);
            
        }
    }
}
