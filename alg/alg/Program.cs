using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Diagnostics; 
//file content  file1_file2       link1_link2        percentage1_percentage2_lineMatches
//graph         node(key)         neighbor           linknode1_linknode2_percentage_lineMatches
//MST          file1_file2        link1_link2        lineMatches_averageSimilarity

public class DSU
{
    public int[] parent;
    public int[] rank;
    public DSU(int size)
    {
        parent = new int[size];
        rank = new int[size];
        for (int i = 0; i < size; i++)
        {
            parent[i] = i;
            rank[i] = 0;
        }
    }
    public int find(int x)
    {
        if (parent[x] != x)
            parent[x] = find(parent[x]);
        return parent[x];
    }
    public void union(int x, int y)
    {
        int rootx = find(x), rooty = find(y);
        if (rootx != rooty)
        {
            if (rank[rootx] < rank[rooty])
                parent[rootx] = rooty;
            else if (rank[rootx] > rank[rooty])
                parent[rooty] = rootx;
            else if (rank[rootx] == rank[rooty])
            {
                parent[rooty] = rootx;
                rank[rootx]++;
            }
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        string filePath = @"C:\\Users\\DELL\\Downloads\\Complete\\Complete\\Easy\\1-Input.xlsx";
        string linkType = "html";
        Stopwatch stopwatch = new Stopwatch();
        Stopwatch stattime = new Stopwatch();
        Stopwatch msttime = new Stopwatch();

        stopwatch.Start();
        FileInfo fileInfo = new FileInfo(filePath);
        //    //3mlt el part bta3 file1_fil2_percentage as when constract graph if its links it will not be true constraction  
        //    //save el percntage 3shan el mst    //da hyl8y function el match aslan
        Dictionary<string, (string, string)> fileContent = new Dictionary<string, (string, string)>();
        //    // the key is two files name as old and the max percentage between 2 edges

        fileContent = ReadFile(filePath, fileInfo, linkType);
        //constract graph

        Dictionary<string, Dictionary<string, string>> graph = ConstructAdjacencyList(fileContent);
        msttime.Start();
        List<KeyValuePair<string, double>> maxSpanningTree = MaximumSpanningTree(graph);
        // List to store the sorted MST edges with average similarity and group index
        List<Tuple<List<string>, double, int>> GroupingResult = FindGroupsStat(graph);

        List<(string, double, double, int)> sortedMSTEdges = SortedMSTByAVG(graph, maxSpanningTree, GroupingResult);
        List<List<(string, double, double, int)>> finalMst2 = SortedMSTByLM2(sortedMSTEdges);
        CreateFile2(finalMst2, @"C:\\Users\\DELL\\Downloads\\Complete\\Complete\\Easy\\ov.xlsx", linkType);
        // break;
        msttime.Stop();
        stattime.Start();

        List<Tuple<List<string>, double, int>> GroupingResult2 = FindGroupsStat(graph);

        CreateGroupingFile(GroupingResult2, @"C:\\Users\\DELL\\Downloads\\Complete\\Complete\\Easy\\zz.xlsx", linkType);
        stattime.Stop();
        long elapsedTimeMillisecondsforstat = stattime.ElapsedMilliseconds;

        long elapsedTimeMillisecondsformst = msttime.ElapsedMilliseconds;

        stopwatch.Stop();

        long elapsedTimeMilliseconds = stopwatch.ElapsedMilliseconds;

        Console.WriteLine($"Total Execution Time: {elapsedTimeMilliseconds} milliseconds");
        Console.WriteLine($"Execution Time for stat: {elapsedTimeMillisecondsforstat} milliseconds");
        Console.WriteLine($"Execution Time for mst: {elapsedTimeMillisecondsformst} milliseconds");
    }
    

    #region create files
    // writing mst
    static void CreateFile2(List<List<(string, double, double, int)>> finalMst, string filePath, String linkType)
    {
        using (ExcelPackage package = new ExcelPackage())
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("sheet1");
            worksheet.Cells[1, 1].Value = "File 1";
            worksheet.Cells[1, 2].Value = "File 2";
            worksheet.Cells[1, 3].Value = "Line Matches";
            int row = 2;
            if (linkType == "excel" || linkType == "html")
            {
                foreach (var sublist in finalMst)
                {
                    foreach (var tuple in sublist)
                    {
                        // Split the first item into two links
                        string[] links = tuple.Item1.Split("_");
                        if (links.Length == 2)
                        {
                            string link1 = links[0].Replace(" ", "");
                            string link2 = links[1].Replace(" ", "");
                            var cell1 = worksheet.Cells[row, 1];
                            cell1.Hyperlink = new Uri(link1);
                            cell1.Value = link1;
                            cell1.Style.Font.Color.SetColor(System.Drawing.Color.Blue); // Set font color to blue
                            cell1.Style.Font.UnderLine = true; // Underline the link
                            var cell2 = worksheet.Cells[row, 2];
                            cell2.Hyperlink = new Uri(link2);
                            cell2.Value = link2;
                            cell2.Style.Font.Color.SetColor(System.Drawing.Color.Blue); // Set font color to blue
                            cell2.Style.Font.UnderLine = true; // Underline the link
                        }
                        worksheet.Cells[row, 3].Value = tuple.Item2;
                        row++;
                    }
                }
            }
            else
            {
                foreach (var sublist in finalMst)
                {
                    foreach (var tuple in sublist)
                    {
                        // Split the first item into two links
                        string[] links = tuple.Item1.Split("_");
                        if (links.Length == 2)
                        {
                            string link1 = links[0].Replace(" ", "");
                            string link2 = links[1].Replace(" ", "");
                            var cell1 = worksheet.Cells[row, 1];
                            var cell2 = worksheet.Cells[row, 2];
                            worksheet.Cells[row, 1].Value = link1;
                            worksheet.Cells[row, 2].Value = link2;
                        }
                        worksheet.Cells[row, 3].Value = tuple.Item2;
                        row++;
                    }
                }
            }
            // Adjust column widths to fit content
            worksheet.Cells.AutoFitColumns();
            FileInfo fileInfo = new FileInfo(filePath);
            package.SaveAs(fileInfo);
        }
    } 
    static void CreateGroupingFile(List<Tuple<List<string>, double, int>> data, string filePath, string linktype)
    {
        using (ExcelPackage package = new ExcelPackage())
        {
          
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("sheet1");

            // Sort data by average similarity in descending order
            data = data.OrderByDescending(x => x.Item2).ToList();

            // Set the header
            worksheet.Cells[1, 1].Value = "Component Index";
            worksheet.Cells[1, 2].Value = "Vertices";
            worksheet.Cells[1, 3].Value = "Average Similarity";
            worksheet.Cells[1, 4].Value = "Component Count";

            // Ensure the first tuple's vertices are sorted by integer value
            var firstTuple = data[0];

            data[0] = Tuple.Create(firstTuple.Item1.OrderBy(i => int.Parse(i)).ToList(), firstTuple.Item2, firstTuple.Item3);

            int row = 2;
            foreach (var x in data)
            {
                worksheet.Cells[row, 1].Value = row - 1;

                // Use StringBuilder for efficient string concatenation
                var verticesBuilder = new StringBuilder();
                foreach (var vertex in x.Item1)
                {
                    if (verticesBuilder.Length > 0)
                        verticesBuilder.Append(", ");
                    verticesBuilder.Append(vertex);
                }

                worksheet.Cells[row, 2].Value = verticesBuilder.ToString();
                worksheet.Cells[row, 3].Value = x.Item2;
                worksheet.Cells[row, 4].Value = x.Item1.Count;
                row++;
            }

            // Save the Excel file
            FileInfo fileInfo = new FileInfo(filePath);
            package.SaveAs(fileInfo);        
        }
    }

    #endregion

    #region Read file-Helper Functions
    private static Dictionary<string, (string, string)> ReadFile(string filePath, FileInfo fileInfo, string linkType)
    {
        Dictionary<string, (string, string)> pairDictionary = new Dictionary<string, (string, string)>();

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;
            // Iterate through each row
            for (int row = 2; row <= rowCount; row++) // Assuming row 1 is header
            {
                // Read values from columns A, B, and C
                string columnAValue = worksheet.Cells[row, 1].Text;
                string columnBValue = worksheet.Cells[row, 2].Text;
                string columnCText = worksheet.Cells[row, 3].Text;
                // Remove parentheses and percentage symbols from the Line Matches value
                string cleanedColumnCText = columnCText.Replace("(", "").Replace(")", "").Replace("%", "");
                // Parse the cleaned text to a double
                double columnCValue = double.Parse(cleanedColumnCText);
                // Store values in the dictionary
                string DECkey = KeyToStore(columnAValue, columnBValue, linkType);
                Match matchA = Regex.Match(columnAValue, @"(\d+(\.\d+)?)%");
                Match matchB = Regex.Match(columnBValue, @"(\d+(\.\d+)?)%");
                var links = $"{columnAValue}_{columnBValue}";
                string twoPercentages = $"{double.Parse(matchA.Groups[1].Value)}_{double.Parse(matchB.Groups[1].Value)}_{columnCValue}";
                pairDictionary[DECkey] = (links, twoPercentages);
                Console.WriteLine("dic : " +pairDictionary[DECkey]);
            }
        }
        return pairDictionary;
    }

    private static string KeyToStore(string link1, string link2, string LinkType)
    {
        string[] file1Parts = link1.Split('/');
        string file1Name = "";
        double file1Percentage = 0;

        string[] file2Parts = link2.Split('/');
        string file2Name = "";
        double file2Percentage = 0;
        if (LinkType == "excel")
        {
            Match match1 = Regex.Match(file1Parts[file1Parts.Length - 1], @"(\d+(\.\d+)?)%");
            Match match2 = Regex.Match(file2Parts[file2Parts.Length - 1], @"(\d+(\.\d+)?)%");
            file1Name = file1Parts[file1Parts.Length - 2];
            file1Percentage = double.Parse(match1.Groups[1].Value);

            file2Name = file2Parts[file2Parts.Length - 2];
            file2Percentage = double.Parse(match2.Groups[1].Value);
        }
        else
        {
            file1Name = file1Parts[file1Parts.Length - 2].Replace("file", "");

            Match match1 = Regex.Match(file1Parts[file1Parts.Length - 1], @"(\d+(\.\d+)?)%");
            Match match2 = Regex.Match(file2Parts[file2Parts.Length - 1], @"(\d+(\.\d+)?)%");
            file1Percentage = double.Parse(match1.Groups[1].Value);

            file2Name = file2Parts[file2Parts.Length - 2].Replace("file", "");
            file2Percentage = double.Parse(match2.Groups[1].Value);
        }
        string key = $"{file1Name}_{file2Name}";
        Console.WriteLine("key value : "+key);
        return key;
    }

    private static double ParsePercentage(string input)
    {
        Match match = Regex.Match(input, @"(\d+(\.\d+)?)%");
       
        return double.Parse(match.Groups[1].Value);
    }
    #endregion

    #region Constract Graph
    static Dictionary<string, Dictionary<string, string>> ConstructAdjacencyList(Dictionary<string, (string, string)> fileContent)
    {
        Dictionary<string, Dictionary<string, string>> adjacencyList = new Dictionary<string, Dictionary<string, string>>();

        foreach (var element in fileContent)
        {
            string[] nodes = element.Key.Split('_'); // Get the nodes in the current pair by splitting them 
            string[] percentages = element.Value.Item2.Split('_');
            // Ensure both nodes are present in the adjacency list
            if (!adjacencyList.ContainsKey(nodes[0]))
            {
                adjacencyList[nodes[0]] = new Dictionary<string, string>();
            }
            if (!adjacencyList.ContainsKey(nodes[1]))
            {
                adjacencyList[nodes[1]] = new Dictionary<string, string>();
            }
            string linkAdded1 = $"{element.Value.Item1}_{percentages[0]}_{percentages[2]}";
            string linkAdded2 = $"{element.Value.Item1}_{percentages[1]}_{percentages[2]}";
            Console.WriteLine( "link 1 : " +linkAdded1 +" link 2 : " + linkAdded2);
            // Add each node to the neighbor list of the other node with the corresponding similarity value
            Console.WriteLine("Node 0 : " + nodes[0] + " Node 1: " + nodes[1]);
            adjacencyList[nodes[0]].Add(nodes[1], linkAdded1);
            adjacencyList[nodes[1]].Add(nodes[0], linkAdded2);
        }
        return adjacencyList;
    }
    #endregion

    #region Grouping part

    public static List<Tuple<List<string>, double, int>> FindGroupsStat(Dictionary<string, Dictionary<string, string>> graph)
    {
        List<Tuple<List<string>, double, int>> stat = new List<Tuple<List<string>, double, int>>();
        HashSet<string> visited = new HashSet<string>();
        int groupIndex = 1;

        foreach (var node in graph.Keys)
        {
            if (visited.Contains(node) == false)
            {
                List<string> group = new List<string>();
                double totalsim = 0, sum = 0, numadj = 0, avgsim = 0;
                DFS(graph, node, visited, group, ref totalsim, ref numadj);
                //BFS(graph, node, visited, group, ref totalsim, ref numadj);
                sum += totalsim;
                if (numadj > 0)
                {
                    avgsim = sum / numadj;
                }
                Tuple<List<string>, double, int> newgroup = new Tuple<List<string>, double, int>(group, Math.Round(avgsim, 1), groupIndex++);
                stat.Add(newgroup);
            }
        }
        return stat;
    }

    public static void DFS(Dictionary<string, Dictionary<string, string>> graph, string node, HashSet<string> visited, List<string> group, ref double totalsim, ref double numadj)
    {
        visited.Add(node);
        group.Add(node);
        foreach (var neigbour in graph[node])
        {
            string name = neigbour.Key;
            string[] LinkParts = neigbour.Value.Split('_');
            double weight = Double.Parse(LinkParts[2]);
            totalsim += weight;
            if (graph.ContainsKey(name) && graph[name].ContainsKey(node))
            {
                numadj++;
            }
            if (visited.Contains(name) == false)
            {
                DFS(graph, name, visited, group, ref totalsim, ref numadj);
            }
        }

    }
    #endregion

    #region MST part
    //1    link1_link2_percentage
    public static List<KeyValuePair<string, double>> MaximumSpanningTree(Dictionary<string, Dictionary<string, string>> graph)
    {
        // Initialize a list to store the edges of the MST
        List<KeyValuePair<string, double>> mstEdges = new List<KeyValuePair<string, double>>();

        // Sort all edges by weight in non-increasing order
        List<KeyValuePair<string, (string, string)>> edges = new List<KeyValuePair<string, (string, string)>>();
        foreach (var nodePair in graph)
        {

            foreach (var neighbor in nodePair.Value)
            {

                edges.Add(new KeyValuePair<string, (string, string)>(nodePair.Key + "_" + neighbor.Key, (neighbor.Key, neighbor.Value)));
            }
        }

        edges.Sort((x, y) =>
        {

            // Splitting the values after '_'
            string[] xParts = x.Value.Item2.Split('_');
            string[] yParts = y.Value.Item2.Split('_');

            // Extracting weight and line matches
            double xWeight = double.Parse(xParts[2]);
            double yWeight = double.Parse(yParts[2]);
            int xLineMatches = int.Parse(xParts[3]);
            int yLineMatches = int.Parse(yParts[3]);


            // Comparing based on weight
            int weightComparison = yWeight.CompareTo(xWeight);

            // If weights are equal, compare based on line matches
            if (weightComparison == 0)
            {
                return yLineMatches.CompareTo(xLineMatches);
            }

            // Otherwise, return the weight comparison
            return weightComparison;
        });

        // Initialize disjoint-set data structure
        DSU dsu = new DSU(graph.Count);
        Dictionary<string, int> nodeIndices = new Dictionary<string, int>();
        int index = 0;
        foreach (var node in graph.Keys)
        {
            nodeIndices[node] = index++;
        }

        // Iterate through sorted edges
        foreach (var edge in edges)
        {
            string[] nodes = edge.Key.Split('_');
            string source = nodes[0];
            string destination = nodes[1];
            string[] link1_link2_percentage_matchingLines = edge.Value.Item2.Split('_');
            double weight = Double.Parse(link1_link2_percentage_matchingLines[3]);


            // Check if adding the edge creates a cycle
            int rootSource = dsu.find(nodeIndices[source]);

            int rootDestination = dsu.find(nodeIndices[destination]);
            string finalLinks = $"{edge.Key}_{link1_link2_percentage_matchingLines[0]}_{link1_link2_percentage_matchingLines[1]}";
            if (rootSource != rootDestination)
            {
                // Add the edge to the MST
                mstEdges.Add(new KeyValuePair<string, double>(finalLinks, Double.Parse(link1_link2_percentage_matchingLines[3])));
                dsu.union(rootSource, rootDestination);
            }
        }



        return mstEdges;
    }

    public static List<(string, double, double, int)> SortedMSTByAVG(Dictionary<string, Dictionary<string, string>> graph, List<KeyValuePair<string, double>> mstEdges, List<Tuple<List<string>, double, int>> groupsStat)
    {
        // Obtain groups with average similarities and collect all nodes in one HashSet
        HashSet<string> groupNodes = new HashSet<string>();
        // Initialize dictionary to store the average similarity of each node
        Dictionary<string, double> nodeAverageSimilarity = new Dictionary<string, double>();
        // Initialize a dictionary to store the group index of each node
        Dictionary<string, int> nodeGroupIndex = new Dictionary<string, int>();

        foreach (var group in groupsStat)
        {
            double avgSimilarity = group.Item2;
            int groupIndex = group.Item3;

            foreach (var node in group.Item1)
            {
                // Store node average similarity
                nodeAverageSimilarity[node] = avgSimilarity;
                nodeGroupIndex[node] = groupIndex;

                // Add node to HashSet
                groupNodes.Add(node);
            }
        }

        // Initialize a list to store the edges of the MST with associated average similarities and group index
        List<(string, double, double, int)> mstEdgesWithAvgSimilarityAndGroupIndex = new List<(string, double, double, int)>();

        // Iterate through sorted edges
        foreach (var edge in mstEdges)
        {
            // Extract file names from URLs
            string[] parts = edge.Key.Split('_');
            string fileName1 = parts[0];
            string fileName2 = parts[1];
            string finalLinks = $"{parts[2]}_{parts[3]}";
            // Check if both file names are present in any group
            double avgSimilarity = groupNodes.Contains(fileName1) && groupNodes.Contains(fileName2)
                ? nodeAverageSimilarity[fileName1]
                : 0; // Default value if not present in any group

            int groupIndex = groupNodes.Contains(fileName1) && groupNodes.Contains(fileName2)
                ? nodeGroupIndex[fileName1] // Use any of the nodes' group index since they should be the same
                : 0; // Default value if not present in any group

            // Add the line matches, average similarity, and group index of source and destination nodes to the list
            //links    line matches    average similarity    group index
            mstEdgesWithAvgSimilarityAndGroupIndex.Add((finalLinks, edge.Value, avgSimilarity, groupIndex));
        }
        
        mstEdgesWithAvgSimilarityAndGroupIndex.Sort((x, y) => y.Item3.CompareTo(x.Item3));
        
        return mstEdgesWithAvgSimilarityAndGroupIndex;
    }

    public static List<List<(string, double, double, int)>> SortedMSTByLM2(List<(string, double, double, int)> mstEdgesByGroup)
    {
        // Group the input list by the last element (group index)
        var groupedByGroupIndex = mstEdgesByGroup.GroupBy(edge => edge.Item4);

        // Initialize a list to store sorted groups
        var sortedGroups = new List<List<(string, double, double, int)>>();

        // Sort each group by line matches (LM) in descending order and convert group index to char
        foreach (var group in groupedByGroupIndex)
        {
            var sortedGroup = group.OrderByDescending(edge => edge.Item2)
                                   .Select(edge => (edge.Item1, edge.Item2, edge.Item3, edge.Item4))
                                   .ToList();
            sortedGroups.Add(sortedGroup);
        }

        return sortedGroups;
    }
    #endregion

}