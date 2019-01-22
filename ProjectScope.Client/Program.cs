using System;
using XMindAPI.LIB;
using LinqToExcel;
using ProjectScope.Client.Models;
using System.Linq;
using System.Collections.Generic;


namespace ProjectScope.Client
{
    class Program
    {
        private static string _defaultXMindSheetName = "projectScope";
        private static string jiraBaseUrl = @"https://nlepsd.atlassian.net/browse/";
        static void Main(string[] args)
        {
            string fileExtension = "xmind";
            string defaultExcelSheetName = "Sprint planning phase 2";
            string dataSource = @"C:\Users\HYS\Downloads\internal_actionlist.xlsx";
            XMindWorkBook book = new XMindWorkBook($"test.{fileExtension}");
            var userStories = ReadDataFromExceFile(dataSource, defaultExcelSheetName);
            string sheetId = book.AddSheet(_defaultXMindSheetName);
            ConfigureXMindWorkBook(book, sheetId, userStories);
            book.Save();
        }
        public static void ConfigureXMindWorkBook(XMindWorkBook book, string sheetId, IDictionary<string, List<UserStory>> data)
        {
            string centralTopicId = book.AddCentralTopic(sheetId, "Phase 2 Scope", XMindStructure.Map);
            foreach (KeyValuePair<string, List<UserStory>> entry in data)
            {
                var currentTopicId = book.AddTopic(centralTopicId, entry.Key);
                entry.Value.ForEach(userStory =>
                    {
                        var userStoryTopicId = book.AddTopic(currentTopicId, $"{userStory.Reference}: {userStory.Name}");
                        book.AddUserTag(userStoryTopicId, userStory.Reference, userStory.Reference);
                        // add label
                        book.AddLabel(userStoryTopicId, $"{jiraBaseUrl}{userStory.Reference}");
                        // add comments
                        if (!String.IsNullOrEmpty(userStory.Comments))
                        {
                            foreach (var item in userStory.Comments.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries))
                            {
                                book.AddTopic(userStoryTopicId, item);
                            }
                        }
                    }
                );
            }
            foreach (UserStory story in data.Values.SelectMany(l => l.AsEnumerable()))
            {
                if (!String.IsNullOrEmpty(story.DependsOn))
                {
                    foreach(var parent in story.DependsOn.Split(','))
                    {
                        book.AddTopicLink(
                            book.GetTopicIdsByUserTagValue(story.Reference, story.Reference).FirstOrDefault(),
                            book.GetTopicIdsByUserTagValue(parent, parent).FirstOrDefault()
                        );
                    }
                }
            }
        }

        public static IDictionary<string, List<UserStory>> ReadDataFromExceFile(string fileName, string sheetName)
        {
            var excel = new ExcelQueryFactory(fileName);
            //excel.AddMapping<UserStory>(item => item.Reference, "ZEB nr");
            //foreach(var item in excel.GetColumnNames(sheetName))
            //{
            //    Console.WriteLine($"column: {item}");
            //}

            var userStories = excel.Worksheet<UserStory>(sheetName)
                .Where(userStory => userStory.Reference != String.Empty).ToList<UserStory>();
            var groupedUserStories = userStories.GroupBy(story => story.Scope);
            return groupedUserStories.ToDictionary(group => group.Key, group => group.ToList());
        }
    }
}
