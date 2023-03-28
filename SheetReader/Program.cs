namespace SheetReader;

internal class Program
{
    private static void Main(string[] args)
    {
        var path = args.Count() > 0 ? args[0] : "../sheets/";
        if (Directory.Exists(path))
        {
            var files = Directory.EnumerateFiles(path);
            
            var subawards = new Dictionary<String, int>();
            foreach (var file in files)
            {
                Console.WriteLine(file);
                var current_sheet = new Sheet(file);
                
                
                foreach(var recipient in current_sheet.subawards.Keys)
                {
                    Console.WriteLine(recipient);
                    if (subawards.ContainsKey(recipient))
                    {
                        subawards[recipient] += current_sheet.subawards[recipient];
                    }else
                    {
                        subawards.Add(recipient, current_sheet.subawards[recipient]);
                    }

                }

            }

            foreach(var recipient in subawards.Keys)
            {
                Console.WriteLine(recipient + ": " + subawards[recipient]);

            }
        }
        else
        {
            throw new FileNotFoundException("Directory " + path + "not found");
        }

    }
}