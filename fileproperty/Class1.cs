// https://stackoverflow.com/questions/5337683/how-to-set-extended-file-properties?noredirect=1&lq=1

using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace fileproperty
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"E:\Projects\Github\MicroSoft\C#\CSharp\demos\example.docx";

            var file = ShellFile.FromFilePath(filePath);

            // Read and Write:

            string[] oldAuthors = file.Properties.System.Author.Value;
            string oldTitle = file.Properties.System.Title.Value;

            file.Properties.System.Author.Value = new string[] { "Author #1", "Author #2" };
            file.Properties.System.Title.Value = "Example Title";

            // Alternate way to Write:

            ShellPropertyWriter propertyWriter = file.Properties.GetPropertyWriter();
            propertyWriter.WriteProperty(SystemProperties.System.Author, new string[] { "Author" });

            propertyWriter.Close();
        }

        /*
         public static void Main(string[] args)
        {
            List<string> arrHeaders = new List<string>();

            Shell32.Shell shell = new Shell32.Shell();
            Shell32.Folder objFolder;

            objFolder = shell.NameSpace(@"C:\temp\testprop");

            for( int i = 0; i < short.MaxValue; i++ )
            {
                string header = objFolder.GetDetailsOf(null, i);
                if (String.IsNullOrEmpty(header))
                    break;
                arrHeaders.Add(header);
            }

            foreach(Shell32.FolderItem2 item in objFolder.Items())
            {
                for (int i = 0; i < arrHeaders.Count; i++)
                {
                    Console.WriteLine(
                      $"{i}\t{arrHeaders[i]}: {objFolder.GetDetailsOf(item, i)}");
                }
            }
        }
         */
    }
}




