using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Xceed.Words.NET;
using Xceed.Document.NET;
using Xceed.Drawing;
using System.Xml.Linq;
using System.Diagnostics;

namespace BibleCompiler2
{
    public partial class Form1 : Form
    {
        static string tkj = "TBC";
        List<Questions> questions = new List<Questions>();
        const string FILENAME2 = "QuestionType.txt";
        string outputPath;
        static string inputDataPath = "Data Files";
        string competitionDocName;
        string inputPath;
        string inputPath2 = Path.Combine(inputDataPath, FILENAME2);
        static string docName = @"Competition Study Guide " + tkj + ".docx";
        string output = "Output Files";
        bool btnInputClicked = false;
        bool btnOutputClicked = false;
        DocX document = DocX.Create(docName, DocumentTypes.Document);
        string font = "Arial";
        int fontSize = 12;
        int spaceFontSize = 4;
        int alnum = 0;
        int maxTComp = 0;
        int maxKComp = 0;
        List<bool> f2True = new List<bool>();
        float margin = 36f; // 72 = 1 inch
        List<Questions> questionsActiveList = new List<Questions>();
        string compNumber = "0";
        Dictionary<string, string> qTypeDict = new Dictionary<string, string>();
        List<string> qTypeList = new List<string>();
        List<string> verseCount = new List<string>();
        List<string> compOrderList;
        List<string> maxCompQuest = new List<string>();
        int selectedCompetitionInt;

        List<string> compSeed = new List<string>();
        List<Questions> quarList = new List<Questions>();
        List<int> mList = new List<int>();
        //List<string> compSeed = new List<string>();
        string filePrefix = "";
        private HashSet<string> printedCompetitions = new HashSet<string>();

        public Form1()
        {
            InitializeComponent();

            // Hard code the input path relative to the debug directory
            string relativePath = @"Data Files\questions.txt";
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string hardCodedPath = Path.Combine(basePath, relativePath);

            if (File.Exists(hardCodedPath))
            {
                inputPath = hardCodedPath;
                lblInputfilepath.Text = inputPath;
                btnInputClicked = true;
            }
            else
            {
                string currentDirectory = Directory.GetCurrentDirectory();
                MessageBox.Show($"The file does not exist at the following path:\n{hardCodedPath}\n\nCurrent Directory:\n{currentDirectory}", 
                                "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            // Set the output path relative to the debug directory
            outputPath = Path.Combine(basePath, "Output Files");
            lblOutputfilepath.Text = outputPath;
        }
        private void standardFormSetup(Button btnAccept, Button btnCancel)
        {
            this.Location = new System.Drawing.Point((Screen.PrimaryScreen.WorkingArea.Width - this.Width) / 2, (Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            AcceptButton = btnSubmit;
            CancelButton = btnExit;
            //this.WindowState = FormWindowState.Maximized;
            pnlCenter.Location = new System.Drawing.Point(this.Width / 2 - pnlCenter.Width / 2, this.Height / 2 - pnlCenter.Height / 2);
            MinimumSize = Size;

        }
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            pnlCenter.Location = new System.Drawing.Point(this.Width / 2 - pnlCenter.Width / 2, this.Height / 2 - pnlCenter.Height / 2);
        }
        private void readQuestionType()
        {
            StreamReader inputFile; //begin reading file
            inputFile = File.OpenText(inputPath2);
            string stringdump = inputFile.ReadLine();
            while (!inputFile.EndOfStream)
            {
                string P = inputFile.ReadLine();
                string[] Q = P.Split('\t');
                qTypeDict.Add(Q[0], Q[1]);
                qTypeList.Add(Q[0]);
            }

        }
        private void createDirectories(string outputFilePath)
        {
            string newPath = Path.Combine(outputFilePath, outputPath);
            Directory.CreateDirectory(newPath);
            Directory.CreateDirectory(Path.Combine(newPath, filePrefix + " Competition Study Guides"));
            Directory.CreateDirectory(Path.Combine(newPath, filePrefix + " Competition Forms"));
        }
        private void pageBreak()
        {
            document.InsertParagraph().Append("").Font(font).FontSize(1).KeepLinesTogether().InsertPageBreakAfterSelf();
        }
        private void studyGuideTableBorderSetup(Table table, int r)
        {
            Border b = new Border(Xceed.Document.NET.BorderStyle.Tcbs_thick, BorderSize.seven, 0, Xceed.Drawing.Color.Black);
            for (int i = 0; i < 3; i++)
            {
                table.Rows[r].Cells[i].FillColor = Xceed.Drawing.Color.LightGray;
                table.Rows[r].Cells[i].SetBorder(TableCellBorderType.Top, b);
                table.Rows[r].Cells[i].SetBorder(TableCellBorderType.Bottom, b);
                table.Rows[r].Cells[i].SetBorder(TableCellBorderType.Left, b);
                table.Rows[r].Cells[i].SetBorder(TableCellBorderType.Right, b);
            }
        }
        private void studyGuideTableBorderSetupType(Table table, int r)
        {
            Border b = new Border(Xceed.Document.NET.BorderStyle.Tcbs_thick, BorderSize.seven, 0, Xceed.Drawing.Color.Black);
            table.Rows[r].Cells[0].SetBorder(TableCellBorderType.Top, b);
            table.Rows[r].Cells[0].SetBorder(TableCellBorderType.Bottom, b);
            table.Rows[r].Cells[0].SetBorder(TableCellBorderType.Left, b);
            table.Rows[r].Cells[0].SetBorder(TableCellBorderType.Right, b);
        }
        private void studyGuideQuestionByType(string z)
        {
            // colWidth is to change the column width for the header table for the "By Type" tables
            int colWidth = 550;
            int numRows = 1;
            int numCol = 2;
            string prevBook = "";
            string currBook = "";
            document.PageHeight = 792;
            document.PageWidth = 612;
            document.MarginTop = margin;
            document.MarginLeft = margin;
            document.MarginRight = margin;
            document.MarginBottom = margin;


            pageBreak();
            for (int i = 0; i < qTypeList.Count; i++)
            {
                if (qTypeList[i] == "F2")
                {
                    i++;
                }
                Table headerTable = document.AddTable(1, 1);
                headerTable.SetColumnWidth(0, colWidth);
                document.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize);
                studyGuideTableBorderSetupType(headerTable, 0);
                headerTable.Rows[0].Cells[0].Paragraphs[0].Append(qTypeDict[qTypeList[i]] + " Question - " + tkj + " Competition " + z)
                    .Font(font).FontSize(fontSize).KeepWithNextParagraph().Bold().Alignment = Alignment.center;
                document.InsertTable(headerTable);

                for (int j = 0; j < questionsActiveList.Count; j++)
                {
                    bool isMatch = false;
                    if (tkj == "TBC")
                    {
                        if (qTypeList[i] == questionsActiveList[j].type && questionsActiveList[j].competitionTBC == z)
                            isMatch = true;
                    }
                    else if (tkj == "KBC")
                    {
                        if (qTypeList[i] == questionsActiveList[j].type && questionsActiveList[j].competitionKBC == z)
                            isMatch = true;
                    }
                    if (isMatch)
                    {
                        currBook = questionsActiveList[j].book + "\n" + questionsActiveList[j].chapter + ":" + questionsActiveList[j].verse;
                        if (currBook != prevBook)
                        {
                            document.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize);
                        }
                        //Creates the tables by Type
                        Table questionTable = document.AddTable(numRows, numCol);
                        questionTable.SetColumnWidth(0, 60);
                        questionTable.SetColumnWidth(1, 490);
                        questionTable.Rows[0].Cells[0].Paragraphs[0].Append(questionsActiveList[j].book + "\n" + questionsActiveList[j].chapter + ":" + questionsActiveList[j].verse)
                            .Font(font).FontSize(fontSize).Bold().Alignment = Alignment.center;
                        if (questionsActiveList[j].type == "Q")
                        {
                            questionTable.Rows[0].Cells[1].Paragraphs[0].Append("Quote " + questionsActiveList[j].book + " " + questionsActiveList[j].chapter + ":" + questionsActiveList[j].verse)
                                .Font(font).FontSize(fontSize).Bold();
                            questionTable.Rows[0].Cells[1].Paragraphs[0].Append("\n" + questionsActiveList[j].question)
                                .Font(font).FontSize(fontSize).Italic();
                        }
                        else
                        {
                            questionTable.Rows[0].Cells[1].Paragraphs[0].Append(questionsActiveList[j].question)
                                .Font(font).FontSize(fontSize).Bold();
                            questionTable.Rows[0].Cells[1].Paragraphs[0].Append("\n" + questionsActiveList[j].answer)
                                .Font(font).FontSize(fontSize).Italic();
                        }

                        document.InsertTable(questionTable);
                        prevBook = currBook;
                    }
                }
            }
        }

        private void studyGuideTableSetUp(int q)
        {
            try
            {
                //These variable are to change the column sizes for the study guide
                int colOne = 60;
                int colTwo = 40;
                int colThree = 550 - (colOne + colTwo);
                //These variables are to change the number of columns 
                int numRows = 1;
                int numCol = 3;
                //
                Table table = document.AddTable(numRows, numCol);
                table.SetColumnWidth(0, colOne);
                table.SetColumnWidth(1, colTwo);
                table.SetColumnWidth(2, colThree);
                alnum++;

                if (tkj == "TBC")
                {
                    if (compNumber != questionsActiveList[q].competitionTBC)
                    {
                        if (compNumber != "0" && !printedCompetitions.Contains(compNumber))
                        {
                            studyGuideQuestionByType(compNumber);
                            printedCompetitions.Add(compNumber);
                            pageBreak();
                        }
                        if (questionsActiveList[q].book != "Book")
                        {
                            document.InsertParagraph()
                                .Append(filePrefix + " " + tkj + " Competition " + questionsActiveList[q].competitionTBC)
                                .Bold().Font(font).FontSize(fontSize * 2).Alignment = Alignment.center;
                            compNumber = questionsActiveList[q].competitionTBC;
                        }
                    }
                }
                else if (tkj == "KBC")
                {
                    if (compNumber != questionsActiveList[q].competitionKBC)
                    {
                        if (compNumber != "0" && !printedCompetitions.Contains(compNumber))
                        {
                            studyGuideQuestionByType(compNumber);
                            printedCompetitions.Add(compNumber);
                            pageBreak();
                        }
                        if (questionsActiveList[q].book != "Book")
                        {
                            document.InsertParagraph()
                                .Append(filePrefix + " " + tkj + " Competition " + questionsActiveList[q].competitionKBC)
                                .Bold().Font(font).FontSize(fontSize * 2).Alignment = Alignment.center;
                            compNumber = questionsActiveList[q].competitionKBC;
                        }
                    }
                }

                if (questionsActiveList[q].book != "Book")
                {
                    if (questionsActiveList[q].type == "V") // Change "V" to display as "Type"
                    {
                        studyGuideTableBorderSetup(table, 0);
                        table.Rows[0].Cells[1].Paragraphs[0]
                            .Append("Type")
                            .Font(font).FontSize(fontSize)
                            .KeepWithNextParagraph()
                            .Bold().Alignment = Alignment.center;
                    }
                    else
                    {
                        table.Rows[0].Cells[1].Paragraphs[0]
                            .Append(questionsActiveList[q].type)
                            .Font(font).FontSize(fontSize)
                            .KeepWithNextParagraph()
                            .Bold().Alignment = Alignment.center;
                    }
                    table.Rows[0].Cells[0].Paragraphs[0]
                        .Append(questionsActiveList[q].book + "\n" + questionsActiveList[q].chapter + ":" + questionsActiveList[q].verse)
                        .Font(font).FontSize(fontSize)
                        .KeepWithNextParagraph()
                        .Bold().Alignment = Alignment.center;

                    if (questionsActiveList[q].type == "Q") // For Quote type
                    {
                        table.Rows[0].Cells[2].Paragraphs[0]
                            .Append("Quote " + questionsActiveList[q].book + " " + questionsActiveList[q].chapter + ":" + questionsActiveList[q].verse + "\n")
                            .Font(font).FontSize(fontSize)
                            .KeepWithNextParagraph()
                            .Bold().Alignment = Alignment.left;
                        table.Rows[0].Cells[2].Paragraphs[0]
                            .Append(questionsActiveList[q].question)
                            .Font(font).FontSize(fontSize)
                            .Italic().Alignment = Alignment.left;
                    }
                    else if (questionsActiveList[q].type != "V")
                    {
                        table.Rows[0].Cells[2].Paragraphs[0]
                            .Append(questionsActiveList[q].question)
                            .Font(font).FontSize(fontSize)
                            .KeepWithNextParagraph()
                            .Bold().Alignment = Alignment.left;
                        table.Rows[0].Cells[2].Paragraphs[0]
                            .Append("\n" + questionsActiveList[q].answer)
                            .Font(font).FontSize(fontSize)
                            .KeepWithNextParagraph()
                            .Italic().Alignment = Alignment.left;
                    }
                    else
                    {
                        document.InsertParagraph().Append("")
                            .Font(font).FontSize(spaceFontSize);
                        table.Rows[0].Cells[2].Paragraphs[0]
                            .Append(questionsActiveList[q].question)
                            .Font(font).FontSize(fontSize)
                            .KeepWithNextParagraph()
                            .Bold().Alignment = Alignment.left;
                    }

                    document.InsertTable(table);
                }
            }
            catch (System.Collections.Generic.KeyNotFoundException)
            {
                MessageBox.Show(questions[q].book + " " + questions[q].chapter + ":" + questions[q].verse +
                    " has an unknown question type (" + questions[q].type + ") in the file, please correct and run again.",
                    "Unknown question type", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void fillQuestionsActiveList()
        {
            List<string> fileNamePrefix = new List<string>();
            string bcv = "";

            int num = 0;
            questionsActiveList.Clear();
            for (int i = 0; i < questions.Count; i++)
            {

                if (tkj == "TBC")
                {
                    if (int.TryParse(questions[i].competitionTBC, out num) && num > 0)
                    {
                        questionsActiveList.Add(questions[i]);
                        bcv = questions[i].competitionTBC + "\t" + questions[i].book + "\t" + questions[i].chapter + "\t" + questions[i].verse;
                    }
                }
                else if (tkj == "KBC")
                {
                    if (int.TryParse(questions[i].competitionKBC, out num) && num > 0)
                    {
                        questionsActiveList.Add(questions[i]);
                        bcv = questions[i].competitionKBC + "\t" + questions[i].book + "\t" + questions[i].chapter + "\t" + questions[i].verse;
                    }
                }
                if (!fileNamePrefix.Contains(questionsActiveList[questionsActiveList.Count - 1].book))
                {
                    fileNamePrefix.Add(questionsActiveList[questionsActiveList.Count - 1].book);
                    filePrefix = fileNamePrefix[0];

                }
                if (!verseCount.Contains(bcv))
                {
                    verseCount.Add(bcv);
                }

            }
            //createSeed();

            for (int j = 1; j < fileNamePrefix.Count; j++)
            {
                filePrefix += ", " + fileNamePrefix[j];
            }
            //this.Text = verseCount.Count.ToString();
            //this.Text = "Bible Challenge Compiler";

        }


        private void studyGuideCreateDoc()
        {
            try
            {
                document.PageHeight = 792;
                document.PageWidth = 612;
                document.MarginTop = margin;
                document.MarginBottom = margin;

                for (int i = 0; i < questionsActiveList.Count; i++)
                {
                    studyGuideTableSetUp(i);
                }
                // After processing all questions, output the final "questions by type" section if needed.
                if (compNumber != "0" && !printedCompetitions.Contains(compNumber))
                {
                    studyGuideQuestionByType(compNumber);
                    printedCompetitions.Add(compNumber);
                    pageBreak();
                }
                document.Save();
            }
            catch (System.IO.IOException)
            {
                OpenFileException ofe = new OpenFileException();
                ofe.err(Path.GetFileName(docName));
            }
        }

        private void createComp()
        {
            // --- Determine competition guide file and paths ---
            string path = Path.Combine(outputPath, filePrefix + " Competition Study Guides");
            string competitionFile = "";
            if (rdbTbccompetition.Checked)
            { /* ... select TBC file ... */
                if (rdb25.Checked) competitionFile = "tbcCompetitionGuide25.txt";
                else if (rdb20.Checked) competitionFile = "tbcCompetitionGuide20.txt";
                else if (rdb13.Checked) competitionFile = "tbcCompetitionGuide13.txt";
                else if (rdb12.Checked) competitionFile = "tbcCompetitionGuide12.txt";
                else if (rdb10.Checked) competitionFile = "tbcCompetitionGuide10.txt";
            }
            else if (rdbKbccompetition.Checked)
            { /* ... select KBC file ... */
                if (rdb25.Checked) competitionFile = "kbcCompetitionGuide25.txt";
                else if (rdb20.Checked) competitionFile = "kbcCompetitionGuide20.txt";
            }
            string baseFolder = (rdbTbccompetition.Checked)
                ? Path.Combine(Application.StartupPath, "Data Files/Teens/Guides (25, 20, 13, 12, 10)")
                : Path.Combine(Application.StartupPath, "Data Files/Kids/Guides (25, 20)");
            string competitionFilePath = Path.Combine(baseFolder, competitionFile);

            // --- Read Competition Number and Guide File ---
            string selectedCompetitionNumber = "";
            if (rdbC1.Checked) selectedCompetitionNumber = "1";
            else if (rdbC2.Checked) selectedCompetitionNumber = "2";
            else if (rdbC3.Checked) selectedCompetitionNumber = "3";
            else if (rdbC4.Checked) selectedCompetitionNumber = "4";
            else if (rdbC5.Checked) selectedCompetitionNumber = "5";
            else if (rdbC6.Checked) selectedCompetitionNumber = "6";
            selectedCompetitionInt = int.Parse(selectedCompetitionNumber);

            List<string> compNumberList = new List<string>();
            compOrderList = new List<string>();
            List<string> compExtraList = new List<string>();
            if (File.Exists(competitionFilePath))
            {
                foreach (string line in File.ReadAllLines(competitionFilePath))
                {
                    string[] parts = line.Split('\t');
                    compNumberList.Add(parts.Length > 0 ? parts[0].Trim() : "");
                    compOrderList.Add(parts.Length > 1 ? parts[1].Trim() : "");
                    compExtraList.Add(parts.Length > 2 ? parts[2].Trim() : "");
                }
            }
            else
            {
                MessageBox.Show($"Error: Competition guide file not found at {competitionFilePath}", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // --- Populate Main Match Questions ---
            List<List<Questions>> selectedQs = matchList(); // Populates quarList and mList (Crucial this works somewhat)

            // --- Prepare Output Document ---
            competitionDocName = Path.Combine(outputPath, filePrefix + " Competition Forms",
                 filePrefix + " " + tkj + " Competition " + getCompetitionNumberName() + " " + getCompetitionOrderName() + " Questions.docx");
            // Use a 'using' statement
            using (DocX compDocument = DocX.Create(competitionDocName, DocumentTypes.Document))
            {
                compDocument.PageHeight = 792;
                compDocument.PageWidth = 612;
                compDocument.MarginTop = margin;
                compDocument.MarginBottom = margin;
                compDocument.MarginLeft = margin;
                compDocument.MarginRight = margin;

                // --- Initialize tracking for used questions ---
                List<Questions> allQuestionsForSelectedComp = questionsActiveList
                    .Where(q => (tkj == "TBC" ? q.competitionTBC : q.competitionKBC) == selectedCompetitionInt.ToString())
                    .ToList();
                HashSet<Questions> questionsUsedSoFar = new HashSet<Questions>();

                // --- Setup variables needed in the loop ---
                int num = 0; // This will hold the competition size (e.g., 25)
                int.TryParse(getCompetitionOrderName(), out num);
                // compNumber seems unused or misused in the original loop's header? Using getCompetitionNumberName() instead.

                // --- Main Loop (Using User's Structure with Minimal Fixes) ---
                for (int i = 1; i < 5; i++) // Hardcoded 4 iterations, 'i' is Match Number 1-4
                {
                    // Add Match Header
                    Table compHeader = compDocument.AddTable(1, 1);
                    compHeader.SetColumnWidth(0, 550); // Adjusted width
                    compDocument.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize);
                    studyGuideTableBorderSetupType(compHeader, 0);
                    // Use getCompetitionNumberName() for C# and i for Match #
                    compHeader.Rows[0].Cells[0].Paragraphs[0].Append($"{tkj} {filePrefix}: Competition {getCompetitionNumberName()} - Match {i}")
                        .Font(font).FontSize((int)(fontSize * 1.5)).Bold().Alignment = Alignment.center;
                    compDocument.InsertTable(compHeader);

                    // Inner Loop - Problematic, but we work around it for uniqueness tracking
                    int questionsAddedThisLoop = 0; // Track questions conceptually added by the flawed inner loop
                    List<Questions> mainQuestionsThisMatch = new List<Questions>(); // Store questions added by inner loop

                    // Writes the standard questions
                    for (int j = 1; j <= num; j++)
                    {
                        insertQuestionFormattedTable(compDocument, selectedQs[i - 1][j - 1], j.ToString());
                        Console.WriteLine(selectedQs[i - 1][j - 1].ToString());
                    } // End inner loop 'j'
                    pageBreak();

                    // --- EXTRAS LOGIC (Call helper with updated used set) ---
                    List<Questions> matchExtras = selectOrderedExtraQuestions(allQuestionsForSelectedComp, questionsUsedSoFar);

                    if (num < selectedQs[i - 1].Count)
                    {
                        // Insert header and questions
                        insertExtraSubsectionHeader(compDocument, "Extra Questions"); // Place header correctly
                        for (int j = num; j < selectedQs[i - 1].Count; j++)
                        {
                            // Use the CORRECT insert function for extras
                            insertQuestionFormattedTable(compDocument, selectedQs[i - 1][j], "__");
                        }
                        Console.WriteLine($"Match {i}: Added {matchExtras.Count} extra questions.");
                    }
                    else
                    {
                        Console.WriteLine($"Match {i}: No extra questions added.");
                    }


                    // --- PAGE BREAK LOGIC---
                    // Add page break after match 'i' and its extras, unless it's the last one (i=4)
                    if (i < 4) // Loop runs for i = 1, 2, 3, 4. Break after 1, 2, 3.
                    {
                        compDocument.InsertParagraph().InsertPageBreakAfterSelf();
                    }

                } // End outer loop 'i'

                // --- Save Document ---
                try { compDocument.Save(); }
                catch (System.IO.IOException ex)
                { /* ... error handling ... */
                    OpenFileException ofe = new OpenFileException();
                    ofe.err(Path.GetFileName(competitionDocName));
                }

            } // End using compDocument
        } // End createComp

        // Selects 7 extra questions (3G, F, M, R, Q) in a specific order if available and unused.
        private List<Questions> selectOrderedExtraQuestions(List<Questions> allCompetitionQuestions, HashSet<Questions> currentlyUsedSet)
        {
            // List to hold the final ordered extra questions for this match.
            List<Questions> orderedExtras = new List<Questions>();
            // Random number generator for selecting among available candidates of a type.
            var random = new Random();
            // Define the required sequence of question types.
            var requiredTypes = new List<string> { "G", "G", "G", "F", "R", "M", "Q" };

            // Filter out used questions and group remaining questions by type.
            // Store them in lists within a dictionary. Shuffle the list for each type initially.
            // Requires Questions class to have Equals/GetHashCode implemented correctly!
            var availableByType = allCompetitionQuestions
                .Where(q => !currentlyUsedSet.Contains(q))
                .GroupBy(q => q.type)
                .ToDictionary(g => g.Key, g => g.OrderBy(q => random.Next()).ToList());

            // Attempt to select one question for each required type in the specified order.
            foreach (string typeNeeded in requiredTypes)
            {
                // Check if the required type exists in our available pool AND has questions left in its list.
                if (availableByType.TryGetValue(typeNeeded, out var candidates) && candidates.Count > 0)
                {
                    // Take the first available question from the (already shuffled) list for this type.
                    Questions selectedQ = candidates[0];
                    // Add it to the results list for this match's extras.
                    orderedExtras.Add(selectedQ);
                    // IMPORTANT: Remove the selected question from the candidates list
                    // to prevent it being selected again immediately (e.g., for the next 'G').
                    candidates.RemoveAt(0);
                }
                else
                {
                    Console.WriteLine($"Warning: Could not find an unused question of type '{typeNeeded}' for extras.");
                }
            }

            // Return the list of selected extra questions in the order they were found.
            return orderedExtras;
        }
        private void createCompetitionTable(DocX docName)
        {
            int numRows = 3;
            int numCol = 2;
            int colOne = 60;
            int colTwo = 490;
            int colThree = 550 - (colOne + colTwo);
            Table compTable = docName.AddTable(numRows, numCol);
            compTable.SetColumnWidth(0, colOne);
            compTable.SetColumnWidth(1, colTwo);
            for (int i = 0; i < quarList.Count; i++)
            {
                if (quarList[i].type == "Q")
                {
                    compTable.Rows[0].MergeCells(0, 2);
                    compTable.Rows[0].Cells[0].Paragraphs[0].Append("Q #" + i.ToString() + " " + "This is a Quote Question. Question!");
                    docName.InsertTable(compTable);
                }
                else if (quarList[i].type == "G")
                {
                    compTable.Rows[0].MergeCells(0, 2);
                    compTable.Rows[0].Cells[0].Paragraphs[0].Append("Q #" + i.ToString() + " " + "This is a General Question. Question!");
                    docName.InsertTable(compTable);
                }
                else if (quarList[i].type == "F")
                {
                    compTable.Rows[0].MergeCells(0, 2);
                    compTable.Rows[0].Cells[0].Paragraphs[0].Append("Q #" + i.ToString() + " " + "This is a Finish This Verse Question. Question!");
                    docName.InsertTable(compTable);
                }
                else if (quarList[i].type == "F2")
                {
                    compTable.Rows[0].MergeCells(0, 2);
                    compTable.Rows[0].Cells[0].Paragraphs[0].Append("Q #" + i.ToString() + " " + "This is a These 2 Verses  Question. Question! Finish These 2 Verses...");
                    docName.InsertTable(compTable);
                }
                else if (quarList[i].type == "M")
                {
                    compTable.Rows[0].MergeCells(0, 2);
                    compTable.Rows[0].Cells[0].Paragraphs[0].Append("Q #" + i.ToString() + " " + "This is a Multiple Answer Question. Question!");
                    docName.InsertTable(compTable);
                }
                else if (quarList[i].type == "R")
                {
                    compTable.Rows[0].MergeCells(0, 2);
                    compTable.Rows[0].Cells[0].Paragraphs[0].Append("Q #" + i.ToString() + " " + "This is a Reference Question. Question!");
                    docName.InsertTable(compTable);
                }
            }

        }
        private void createStudyGuide()
        {
            compNumber = "0";  // Reset the global competition number so sections order correctly.
            printedCompetitions.Clear(); // Reset the printed competitions for each new study guide creation.
            string path = Path.Combine(outputPath, filePrefix + " Competition Study Guides");
            docName = Path.Combine(path, filePrefix + " Competition Study Guide " + tkj + ".docx");
            document = DocX.Create(docName, DocumentTypes.Document);

            studyGuideCreateDoc();
        }

        //This function is for finish the verse
        private string[] finishVerse(string[] v)
        {
            // 0 CompTeen	1 CompKid	2 CompJump	3 Book	4 Chapter	5 Verse	6 Type	7 Question	8 Answer
            string combineWords = "";
            string remainder = "";
            string sq = v[7];
            string[] words = sq.Split(' ');
            string[] ret = new string[v.Length];
            int nw = 4;
            for (int i = 0; i < v.Length; i++)
            {
                ret[i] = v[i];
            }
            ret[6] = "F";
            if (words.Length < 7)
            {
                nw = 2;
            }
            for (int i = 0; i < words.Length; i++)
            {
                //This adds dashes between the first few words
                if (i < nw)
                {
                    combineWords += words[i];
                    combineWords += " - ";
                }
                else if (i == nw)
                {
                    // This adds the dots at the end of the finish the verse question
                    combineWords += words[i];
                    combineWords += "...";
                }
                else // Adds everything that was not a part of the finish the verse question into the answer
                {
                    remainder += (words[i] += " ");
                }

            }
            //returns the broken question and answers
            ret[8] = remainder;
            ret[7] = combineWords;
            return (ret);
        }
        private void f2()
        {
            List<Questions> fq = new List<Questions>();
            for (int i = 0; i < questions.Count; i++)
            {
                if (questions[i].type == "F")
                {
                    fq.Add(questions[i]);
                }
            }
            for (int i = 0; i < questions.Count; i++)
            {
                if (questions[i].type == "F2")
                {
                    for (int k = 0; k < fq.Count; k++)
                    {
                        if (questions[i].book == fq[k].book && questions[i].chapter == fq[k].chapter && questions[i].verse == fq[k].verse)
                        {
                            questions[i].question = fq[k].question;
                            questions[i].answer = fq[k].answer + "* " + fq[k + 1].question.Replace(" - ", " ").Replace("...", " ") + fq[k + 1].answer;
                        }
                    }

                }
            }
            for (int i = 0; i < questions.Count; i++)
            {
                if (questions[i].type == "F3")
                {
                    for (int k = 0; k < fq.Count; k++)
                    {
                        if (questions[i].book == fq[k].book && questions[i].chapter == fq[k].chapter && questions[i].verse == fq[k].verse)
                        {
                            questions[i].question = fq[k].question;
                            questions[i].answer = fq[k].answer + "* " + fq[k + 1].question.Replace(" - ", " ").Replace("...", " ") + fq[k + 1].answer + fq[k + 2].question.Replace(" - ", " ").Replace("...", " ") + fq[k + 2].answer;
                        }
                    }

                }
            }
        }
        private void loadFile(string FILENAME)
        {
            string vq = "";
            string fv = "";
            int rdbNum = 0;
            StreamReader inputFile; //begin reading file
            inputFile = File.OpenText(FILENAME);
            string stringdump = inputFile.ReadLine();
            int count = 0;
            while (!inputFile.EndOfStream)
            {
                try
                {
                    string P = inputFile.ReadLine().Replace("\"", string.Empty).Replace('–', '-').Replace((char)2014, '-').Replace((char)8217, '\'');
                    //string P = inputFile.ReadLine().Replace("\"",string.Empty).Replace('�', '\'');

                    string[] Q = P.Split('\t');
                    for (int i = 0; i < Q.Length; i++)
                    {
                        Q[i] = Q[i].Trim();
                    }
                    count++;
                    if (Q.Length != 9)
                    {
                        MessageBox.Show("Column Error line \n" + Q[3] + " " + Q[4] + " " + Q[5] + " " + Q[6] + "\n The Program will now close", "Line Error", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);


                    }

                    if (Q[6] == "Q")
                    {

                        Q[7] = vq;
                        string[] v = finishVerse(Q);
                        Questions f = new Questions(v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], v[8]);

                        if (int.TryParse(f.competitionTBC, out rdbNum))
                        {
                            maxTComp = rdbNum;
                        }
                        if (int.TryParse(f.competitionKBC, out rdbNum))
                        {
                            maxKComp = rdbNum;
                        }

                        if (questions[questions.Count - 1].type == "F2")
                        {
                            questions.Insert(questions.Count - 1, f);
                        }
                        else
                        {
                            questions.Add(f);
                        }


                    }
                    else if (Q[6] == "V")
                    {
                        fv = Q[7];
                    }
                    int num = 0;
                    Questions q = new Questions(Q[0], Q[1], Q[2], Q[3], Q[4], Q[5], Q[6], Q[7], Q[8]);
                    // This if statement makes it so that the questions are added to the list depending on which button is selected
                    questions.Add(q);

                    if (q.type == "V")
                    {
                        vq = q.question;
                    }

                }
                catch (IndexOutOfRangeException e)
                {
                    Close();
                }
            }

            f2();
            Questions r = new Questions("X", "X", "X", "Book", "Chapter", "Verse", "Type", "Question", "Answer");
            questions.Add(r);
        }
        private void radioButtonCount(RadioButton rdb, int inputNum)
        {
            int rdbNum = 0;
            int.TryParse(rdb.Text, out rdbNum);
            if (rdbNum <= inputNum)
            {
                rdb.Visible = true;
            }
            else
            {
                rdb.Visible = false;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            frmSplash splash = new frmSplash();
            splash.ShowDialog();
            standardFormSetup(btnSubmit, btnExit);
            rdbC1.Checked = true;
            rdb25.Checked = true;
            //createDirectories(".");
            radioBtn();
            readQuestionType();
            pnlDoc.Enabled = false;
            lblInputfilepath.Text = inputPath;
            btnSubmit.Enabled = false;

            // --- wire C1‑C6 to the generic handler ---
            rdbC1.CheckedChanged += rdbCompNumber_CheckedChanged;
            rdbC2.CheckedChanged += rdbCompNumber_CheckedChanged;
            rdbC3.CheckedChanged += rdbCompNumber_CheckedChanged;
            rdbC4.CheckedChanged += rdbCompNumber_CheckedChanged;
            rdbC5.CheckedChanged += rdbCompNumber_CheckedChanged;
            rdbC6.CheckedChanged += rdbCompNumber_CheckedChanged;
        }
        // Helper method to return the selected order name (as a string) for the competition.
        private string getCompetitionOrderName()
        {
            if (rdbTbccompetition.Checked)
            {
                if (rdb25.Checked)
                    return "25";
                else if (rdb20.Checked)
                    return "20";
                else if (rdb13.Checked)
                    return "13";
                else if (rdb12.Checked)
                    return "12";
                else if (rdb10.Checked)
                    return "10";
            }
            else if (rdbKbccompetition.Checked)
            {
                if (rdb25.Checked)
                    return "25";
                else if (rdb20.Checked)
                    return "20";
            }
            return "";
        }

        private string getCompetitionNumberName()
        {
            if (rdbC1.Checked) return "1";
            else if (rdbC2.Checked) return "2";
            else if (rdbC3.Checked) return "3";
            else if (rdbC4.Checked) return "4";
            else if (rdbC5.Checked) return "5";
            else if (rdbC6.Checked) return "6";

            return ""; // Default value if none are selected
        }


        //Submit and Exit Button Functions {
        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void btnSubmit_Click(object sender, EventArgs e)
        {
            radioBtn();
            createDirectories(outputPath);
            // First 3 fill-in
            if (rdbTbcFirst3.Checked || rdbKbcFirst3.Checked)
            {
                createFirst3();
            }
            // Quote Fill-in
            else if (rdbTbcQuoteFill.Checked || rdbKbcQuoteFill.Checked)
            {
                createQuoteFill();
            }
            // Competition
            else if (pnlComp.Visible)
            {
                createComp();
                //createCompetition();

                System.Diagnostics.Process.Start(competitionDocName);
            }
            // Study Guide
            else
            {
                createStudyGuide();
                string studyGuideDocName = Path.Combine(outputPath, filePrefix + " Competition Study Guides", filePrefix + " Competition Study Guide " + tkj + ".docx");

                System.Diagnostics.Process.Start(studyGuideDocName);
            }
            btnSubmit.Enabled = false;
        }

        private void radioBtn()
        {
            questionsActiveList.Clear();
            verseCount.Clear();
            if (rdbTbcstudy.Checked || rdbTbcFirst3.Checked || rdbTbcQuoteFill.Checked || rdbTbccompetition.Checked)
            {
                tkj = "TBC";
                btnSubmit.Enabled = true;
            }
            else if (rdbKbcstudy.Checked || rdbKbcFirst3.Checked || rdbKbcQuoteFill.Checked || rdbKbccompetition.Checked)
            {
                tkj = "KBC";
                btnSubmit.Enabled = true;
            }
            //else if (rdbC1.Checked || rdbC2.Checked || rdbC3.Checked || rdbC4.Checked || rdbC5.Checked || rdbC6.Checked ||
            //         rdb10.Checked || rdb25.Checked)
            //{
            btnSubmit.Enabled = true;
            //}
            fillQuestionsActiveList();
        }

        //These Functions will change what is being added to the Questions Active List
        //Then they change tkj to it's respective string depending on which button is selected

        //Radio Button Functions
        private void rdbTbcstudy_CheckedChanged(object sender, EventArgs e)
        {

            showPanel();
        }

        private void rdbKbcstudy_CheckedChanged(object sender, EventArgs e)
        {
            showPanel();
        }
        private void showPanel()
        {
            pnlComp.Visible = false;
            pnlQuestions.Visible = false;
            radioBtn();
            countTypes();
        }
        private void maxrdbCount(int maxNum)
        {
            radioButtonCount(rdbC1, maxNum);
            radioButtonCount(rdbC2, maxNum);
            radioButtonCount(rdbC3, maxNum);
            radioButtonCount(rdbC4, maxNum);
            radioButtonCount(rdbC5, maxNum);
            radioButtonCount(rdbC6, maxNum);
        }
        // Handles Teens competition radio‑button selection
        private void rdbTbccompetition_CheckedChanged(object sender, EventArgs e)
        {
            if (!rdbTbccompetition.Checked) return;          // ignore un‑check
            pnlComp.Visible = true;                          // show panels
            pnlQuestions.Visible = true;
            rdb10.Visible = rdb12.Visible = rdb13.Visible = true;
            radioBtn();                                      // set tkj = "TBC" & rebuild list
            countTypes();                                    // refresh counts
            maxrdbCount(maxTComp);                           // toggle C# radios
        }
        private void rdbKbccompetition_CheckedChanged(object sender, EventArgs e)
        {
            if (!rdbKbccompetition.Checked) return;
            pnlComp.Visible = true;
            pnlQuestions.Visible = true;
            rdb10.Visible = rdb12.Visible = rdb13.Visible = false;
            radioBtn();                                      // set tkj = "KBC" & rebuild list
            countTypes();                                    // refresh counts
            maxrdbCount(maxKComp);
        }

        private void rdbTbcFirst3_CheckedChanged(object sender, EventArgs e)
        {
            pnlComp.Visible = false;
            pnlQuestions.Visible = false;
            radioBtn();
            countTypes();
        }
        private void rdbTbcQuoteFill_CheckedChanged(object sender, EventArgs e)
        {
            pnlComp.Visible = false;
            pnlQuestions.Visible = false;
            radioBtn();
            countTypes();
        }
        private void rdbKbcFirst3_CheckedChanged(object sender, EventArgs e)
        {
            pnlComp.Visible = false;
            pnlQuestions.Visible = false;
            radioBtn();
            countTypes();
        }
        private void rdbKbcQuoteFill_CheckedChanged(object sender, EventArgs e)
        {
            pnlComp.Visible = false;
            pnlQuestions.Visible = false;
            radioBtn();
            countTypes();
        }


        //Input and Output File Funcitons {
        private void btnInputfile_Click(object sender, EventArgs e)
        {
            rdbTbcstudy.Checked = false;
            DialogResult Inputfile = openFile.ShowDialog();
            if (Inputfile == DialogResult.OK)
            {
                btnInputClicked = true;
                inputPath = openFile.FileName;
                lblInputfilepath.Text = inputPath;
                if (btnOutputClicked == false)
                {
                    btnSubmit.Enabled = false;
                    pnlDoc.Enabled = false;
                }
                else
                {
                    btnSubmit.Enabled = true;
                    pnlDoc.Enabled = true;
                }
            }

            countTypes();
        }

        private void btnOutputfile_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowser.ShowDialog();
            if (result == DialogResult.OK)
            {
                btnOutputClicked = true;
                outputPath = folderBrowser.SelectedPath;
                outputPath = folderBrowser.SelectedPath;
                lblOutputfilepath.Text = outputPath;
                if (btnInputClicked == false)
                {
                    btnSubmit.Enabled = false;
                    pnlDoc.Enabled = false;
                }
                else
                {
                    btnSubmit.Enabled = true;
                    pnlDoc.Enabled = true;
                }
            }

        }
        // Rebuilds the ListBox of type counts, filtering to the
        // selected competition if one is chosen when pnlComp is visible.
        private void countTypes()
        {
            Dictionary<string, int> typeCounts = new Dictionary<string, int>();

            string selComp = getCompetitionNumberName();        // selected comp #
            bool filter = pnlComp.Visible && !string.IsNullOrEmpty(selComp);

            foreach (var q in questionsActiveList)
            {
                // skip if filtering and question isn’t in the chosen competition
                if (filter)
                {
                    bool inComp = (tkj == "TBC") ? q.competitionTBC == selComp
                                                 : q.competitionKBC == selComp;
                    if (!inComp) continue;
                }

                if (typeCounts.ContainsKey(q.type)) typeCounts[q.type]++;
                else typeCounts[q.type] = 1;
            }

            // refresh ListBox with header then counts
            lsbQuestionCount.Items.Clear();
            lsbQuestionCount.Items.Add("");
            lsbQuestionCount.Items.Add(frmt("Question", ""));
            lsbQuestionCount.Items.Add(frmt("Type", "Count"));
            lsbQuestionCount.Items.Add(frmt("--------", "-----"));
            foreach (var tc in typeCounts)
                if (tc.Key != "Type")
                    lsbQuestionCount.Items.Add(frmt(tc.Key, tc.Value.ToString()));
        }
        private string frmt(string s1, string s2)
        {
            return string.Format("  {0, -8}|{1, 5}  ", s1, s2);
        }
        private void createSeed()
        {
            compSeed.Clear();
            this.Text = quarList.Count.ToString();
            //lsbTest.Items.Clear();
            var seed = new Random();
            for (int i = 0; i < verseCount.Count; i++)
            {
                int seednum = seed.Next(0, verseCount.Count);
                if (verseCount[seednum].StartsWith(selectedCompetitionInt.ToString()))
                {
                    if (!compSeed.Contains(verseCount[seednum]))
                    {
                        compSeed.Add(verseCount[seednum]);
                        lsbTest2.Items.Add(verseCount[seednum]);
                    }
                }
                else
                {
                    i--;
                }


            }
            lsbTest2.Items.Add("-------------");

        }

        private List<List<Questions>> matchList()
        {
            List<List<Questions>> selectedQs = new List<List<Questions>>();
            HashSet<string> usedBCV = new HashSet<string>();
            string questionType = "";
            bool found = false;
            int c = selectedCompetitionInt;
            createSeed();
            for (int g = 0; g < 4; g++)
            {
                selectedQs.Add(new List<Questions>());
                //quarList.Clear();
                for (int h = 0; h < compOrderList.Count; h++)
                {
                    bool valid = true;
                    found = false;
                    questionType = compOrderList[h];
                    for (int i = 0; i < compSeed.Count; i++)
                    {
                        //take the first competion number from the first item in the Verse Count list 
                        string[] firstLineCompNum = compSeed[i].Split('\t');
                        if (firstLineCompNum[0] == c.ToString())
                        {
                            string[] flcn = firstLineCompNum;
                            for (int j = 0; j < questionsActiveList.Count; j++)
                            {
                                valid = true;
                                if (quarList.Count > 4)
                                {
                                    int count = quarList.Count - 1;
                                    while (!((count == 0) || count == quarList.Count - 4))
                                    {
                                        if (questionsActiveList[j].book == quarList[count].book && questionsActiveList[j].chapter == quarList[count].chapter && questionsActiveList[j].verse == quarList[count].verse)
                                        {
                                            valid = false;
                                            break;
                                        }
                                        count--;
                                    }
                                }
                                //if statement that matches the book, chapter, verse, and question type with the Active List
                                string bcvKey = questionsActiveList[j].book + "" + questionsActiveList[j].chapter + "_" + questionsActiveList[j].verse;
                                if (valid && !quarList.Contains(questionsActiveList[j]) && !usedBCV.Contains(bcvKey))
                                {
                                    usedBCV.Add(bcvKey);
                                }
                                if (valid && questionsActiveList[j].book == flcn[1] && questionsActiveList[j].chapter == flcn[2] && questionsActiveList[j].verse == flcn[3] && questionsActiveList[j].type == questionType && !quarList.Contains(questionsActiveList[j]))
                                {
                                    //Add questions to the Quarantine List
                                    quarList.Add(questionsActiveList[j]);
                                    mList.Add(g);

                                    string temp = compSeed[i];
                                    compSeed.Remove(compSeed[i]);
                                    compSeed.Add(temp);
                                    //compSeed.Add(string.Join("\t", firstLineCompNum));
                                    lsbTest.Items.Add(h.ToString() + " " + questionsActiveList[j]);
                                    selectedQs[g].Add(questionsActiveList[j]);
                                    //lsbTest2.Items.Add(temp);
                                    temp = "";
                                    found = true;
                                    break;

                                }
                                if (questionType == "F2")
                                {
                                    this.Text = (!found + (questionType == "F2").ToString() + (compSeed.Count - 1 == i).ToString());
                                }

                            }
                            if (found)
                            {

                                //i = 0;
                                break;
                            }


                        }
                        if (valid && !found && compSeed.Count - 1 == i)
                        {
                            if (questionType == "F2")
                            {
                                questionType = "F";
                                i = 0;
                            }
                            else
                            {
                                for (int k = compSeed.Count - 1; k > 0; k--)
                                {
                                    if (mList[k] != g && questionType == quarList[k].type)
                                    {
                                        quarList.RemoveAt(k);
                                        mList.RemoveAt(k);
                                        i = 0;
                                    }
                                }
                            }
                        }
                    }
                }
                lsbTest.Items.Add("--------------------------");
            }
            return selectedQs;
        }

        private void insertDocumentTitleHeader(DocX compDocument, int matchNumber)
        {
            string titleText = $"{tkj} {questionsActiveList[6].book}: Competition {getCompetitionNumberName()} - Match {matchNumber}";
            Table titleTable = compDocument.AddTable(1, 3);
            titleTable.SetWidths(new float[] { 10, 10, 580 });
            studyGuideTableBorderSetup(titleTable, 0);
            titleTable.Rows[0].MergeCells(0, 2);
            titleTable.Rows[0].Cells[0].Paragraphs[0]
                .Append(titleText)
                .Font(font)
                .FontSize(fontSize * 2)
                .Bold()
                .Alignment = Alignment.center;
            compDocument.InsertTable(titleTable);
            compDocument.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize);
        }
        private void insertQuestionFormattedTable(DocX compDocument, Questions q, string questionNumber)
        {
            string headerText = "";
            if (q.type == "F" || q.type == "F2")
            {
                headerText = "Q #" + questionNumber + " is a " + qTypeDict[q.type] + " type of Question. Question! " + qTypeDict[q.type] + "...";
            }
            else
            {
                headerText = "Q #" + questionNumber + " is a " + qTypeDict[q.type] + " Question. Question!";
            }

            Table table = compDocument.AddTable(3, 2);
            table.SetWidths(new float[] { 50, 500 });
            Border b = new Border(Xceed.Document.NET.BorderStyle.Tcbs_thick, BorderSize.seven, 0, Xceed.Drawing.Color.Black);
            Border a = new Border(Xceed.Document.NET.BorderStyle.Tcbs_thick, BorderSize.four, 0, Xceed.Drawing.Color.Black);

            if (questionNumber == "__")
            {
                // For extra questions: set all cells to white.
                for (int r = 0; r < 3; r++)
                {
                    for (int c = 0; c < 2; c++)
                    {
                        table.Rows[r].Cells[c].FillColor = Xceed.Drawing.Color.White;
                        table.Rows[r].Cells[c].SetBorder(TableCellBorderType.Top, b);
                        table.Rows[r].Cells[c].SetBorder(TableCellBorderType.Bottom, b);
                        table.Rows[r].Cells[c].SetBorder(TableCellBorderType.Left, b);
                        table.Rows[r].Cells[c].SetBorder(TableCellBorderType.Right, b);
                    }
                }
            }
            else
            {
                for (int c = 0; c < 2; c++)
                {
                    table.Rows[0].Cells[c].FillColor = Xceed.Drawing.Color.LightGray;
                    table.Rows[0].Cells[c].SetBorder(TableCellBorderType.Top, b);
                    table.Rows[0].Cells[c].SetBorder(TableCellBorderType.Bottom, b);
                    table.Rows[0].Cells[c].SetBorder(TableCellBorderType.Left, b);
                    table.Rows[0].Cells[c].SetBorder(TableCellBorderType.Right, b);
                }
                // Set rows 1 and 2 to White.
                for (int r = 1; r < 3; r++)
                {
                    for (int c = 0; c < 2; c++)
                    {
                        table.Rows[r].Cells[c].FillColor = Xceed.Drawing.Color.White;
                        table.Rows[r].Cells[c].SetBorder(TableCellBorderType.Top, a);
                        table.Rows[r].Cells[c].SetBorder(TableCellBorderType.Bottom, a);
                        table.Rows[r].Cells[c].SetBorder(TableCellBorderType.Left, a);
                        table.Rows[r].Cells[c].SetBorder(TableCellBorderType.Right, a);
                    }
                }
            }

            if (q.type != "Q")
            {
                table.Rows[0].MergeCells(0, 1);
                table.Rows[0].Cells[0].Paragraphs[0].Append(headerText)
                    .Font(font).FontSize(fontSize).KeepWithNextParagraph().Alignment = Alignment.left;
                table.Rows[1].Cells[0].Paragraphs[0].Append("Q" + questionNumber)
                     .Font(font).FontSize(fontSize).Bold().KeepWithNextParagraph().Alignment = Alignment.left;
                table.Rows[1].Cells[1].Paragraphs[0].Append(q.question)
                     .Font(font).FontSize(fontSize).KeepWithNextParagraph().Alignment = Alignment.left;
                string addr = q.chapter + ":" + q.verse;
                table.Rows[2].Cells[0].Paragraphs[0].Append(q.book + "\n")
                     .Font(font).FontSize(7).KeepWithNextParagraph().Alignment = Alignment.left;
                table.Rows[2].Cells[0].Paragraphs[0].Append(addr)
                     .Font(font).FontSize(fontSize).KeepWithNextParagraph().Alignment = Alignment.left;
                table.Rows[2].Cells[1].Paragraphs[0].Append(q.answer)
                     .Font(font).FontSize(fontSize).Italic().KeepWithNextParagraph().Alignment = Alignment.left;
                compDocument.InsertTable(table);
                compDocument.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize);
            }
            else
            {
                table.Rows[0].MergeCells(0, 1);
                table.Rows[0].Cells[0].Paragraphs[0].Append(headerText)
                    .Font(font).FontSize(fontSize).KeepWithNextParagraph().Alignment = Alignment.left;
                table.SetWidths(new float[] { 50, 500 });
                table.Rows[1].Cells[0].Paragraphs[0].Append("Q" + questionNumber)
                     .Font(font).FontSize(fontSize).Bold().KeepWithNextParagraph().Alignment = Alignment.left;
                table.Rows[1].Cells[1].Paragraphs[0].Append("Quote " + q.book + " " + q.chapter + ":" + q.verse)
                     .Font(font).FontSize(fontSize).KeepWithNextParagraph().Alignment = Alignment.left;
                string addr = q.chapter + ":" + q.verse;
                table.Rows[2].Cells[0].Paragraphs[0].Append(q.book + "\n")
                     .Font(font).FontSize(7).KeepWithNextParagraph().Alignment = Alignment.left;
                table.Rows[2].Cells[0].Paragraphs[0].Append(addr)
                     .Font(font).FontSize(fontSize).KeepWithNextParagraph().Alignment = Alignment.left;
                table.Rows[2].Cells[1].Paragraphs[0].Append(q.question)
                     .Font(font).FontSize(fontSize).Italic().KeepWithNextParagraph().Alignment = Alignment.left;
                compDocument.InsertTable(table);
                compDocument.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize);
            }
        }
        private void insertBonusOrNewQuizzerHeader(DocX compDocument, string lineText)
        {
            Table specialTable = compDocument.AddTable(1, 3);
            specialTable.SetWidths(new float[] { 10, 10, 580 });
            studyGuideTableBorderSetup(specialTable, 0);
            specialTable.Rows[0].MergeCells(0, 2);
            specialTable.Rows[0].Cells[0].Paragraphs[0]
                .Append(lineText)
                .Font(font)
                .FontSize(10)
                .Bold()
                .Highlight(Highlight.yellow)
                .KeepWithNextParagraph()
                .Alignment = Alignment.left;
            compDocument.InsertTable(specialTable);
            compDocument.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize);
        }

        private void insertExtraSubsectionHeader(DocX compDocument, string headerText)
        {
            Paragraph header = compDocument.InsertParagraph();

            header.Append(headerText + ":")
                  .Font(font)
                  .FontSize(12)
                  .Bold()
                  .KeepWithNextParagraph()

                  .Alignment = Alignment.center;
            header.SpacingAfter(spaceFontSize);
        }

        private void insertExtraQuestionsHeader(DocX compDocument)
        {
            // Insert a page break before the "EXTRA QUESTIONS" header
            compDocument.InsertParagraph().InsertPageBreakAfterSelf();

            //pageBreak();
            Table extraHeaderTable = compDocument.AddTable(1, 1);
            extraHeaderTable.SetWidths(new float[] { 580 });
            extraHeaderTable.Rows[0].Cells[0].FillColor = Xceed.Drawing.Color.LightGray;
            Border b = new Border(Xceed.Document.NET.BorderStyle.Tcbs_thick, BorderSize.seven, 0, Xceed.Drawing.Color.Black);
            extraHeaderTable.Rows[0].Cells[0].SetBorder(TableCellBorderType.Top, b);
            extraHeaderTable.Rows[0].Cells[0].SetBorder(TableCellBorderType.Bottom, b);
            extraHeaderTable.Rows[0].Cells[0].SetBorder(TableCellBorderType.Left, b);
            extraHeaderTable.Rows[0].Cells[0].SetBorder(TableCellBorderType.Right, b);
            Paragraph header = extraHeaderTable.Rows[0].Cells[0].Paragraphs[0];
            header.Append("EXTRA QUESTIONS").Font(font).FontSize(12).Bold();
            header.Alignment = Alignment.center;
            compDocument.InsertTable(extraHeaderTable);
        }





        private void lblInputfilepath_TextChanged(object sender, EventArgs e)
        {
            questions.Clear();
            loadFile(inputPath);
            if (lblInputfilepath.Text != "")
            {
                pnlDoc.Enabled = true;
                rdbTbcstudy.Checked = true;
            }
            else
            {
                pnlDoc.Enabled = false;
            }
        }
        private void createQuoteFill()
        {
            string path = Path.Combine(outputPath, filePrefix + " Competition Study Guides");
            System.IO.Directory.CreateDirectory(path); // Ensure directory exists

            string docName = Path.Combine(path, "Quote Fill-In Guide " + tkj + ".docx");
            DocX quoteDoc = DocX.Create(docName, DocumentTypes.Document);
            quoteDoc.PageHeight = 792;
            quoteDoc.PageWidth = 612;
            quoteDoc.MarginTop = 36f;
            quoteDoc.MarginBottom = 18f;
            quoteDoc.MarginLeft = 36f;
            quoteDoc.MarginRight = 36f;

            List<Questions> quoteList = questionsActiveList.Where(q => q.type == "Q").ToList();

            bool firstQuote = true;
            Paragraph lastParagraphOfPreviousQuote = null; // Variable to hold the reference

            foreach (var q in quoteList)
            {
                // *** CHANGE 1: Apply page break to the previous quote's last paragraph ***
                if (!firstQuote && lastParagraphOfPreviousQuote != null)
                {
                    // Add the page break AFTER the last paragraph of the PREVIOUS quote
                    lastParagraphOfPreviousQuote.InsertPageBreakAfterSelf();
                }
                firstQuote = false; // Set flag for next iteration



                // Build title (No changes here)
                string competitionNumber = (tkj == "TBC") ? q.competitionTBC : q.competitionKBC;
                string title = $"{tkj} Quote Fill-In – C{competitionNumber} {q.book} {q.chapter}:{q.verse}";
                quoteDoc.InsertParagraph(title)
                    .Font(font).FontSize(18).Bold().Alignment = Alignment.center;
                quoteDoc.InsertParagraph().Append("").Font(font).FontSize(18).LineSpacing = 14f;

                // Difficulty 1 (No changes here)
                string fill1 = maskQuote(q.question, 1);
                quoteDoc.InsertParagraph(fill1)
                    .Font(font).FontSize(18).LineSpacing = 14f;
                quoteDoc.InsertParagraph().Append("").Font(font).FontSize(18).LineSpacing = 14f;

                // Difficulty 2 (No changes here)
                string fill2 = maskQuote(q.question, 2);
                quoteDoc.InsertParagraph(fill2)
                    .Font(font).FontSize(18).LineSpacing = 14f;
                quoteDoc.InsertParagraph().Append("").Font(font).FontSize(18).LineSpacing = 14f;

                // Difficulty 3
                string fill3 = maskQuote(q.question, 3);
                // *** CHANGE 3: Store reference to the paragraph containing fill3 ***
                lastParagraphOfPreviousQuote = quoteDoc.InsertParagraph(fill3); // Store this paragraph
                lastParagraphOfPreviousQuote.Font(font).FontSize(18).LineSpacing = 14f; // Apply formatting to the stored paragraph

                // Extra space after fill3 is already removed from previous step.
            }

            // Save and open (No changes here)
            try
            {
                quoteDoc.Save();
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docName) { UseShellExecute = true });
            }
            catch (System.IO.IOException)
            {
                OpenFileException ofe = new OpenFileException();
                ofe.err(Path.GetFileName(docName));
            }
            // Consider Dispose if not using 'using'
        }

        private void createFirst3()
        {
            string path = Path.Combine(outputPath, filePrefix + " Competition Study Guides");
            string first3DocName = Path.Combine(path, filePrefix + " Competition First3 Guide " + tkj + ".docx");
            DocX first3Doc = DocX.Create(first3DocName, DocumentTypes.Document);
            first3Doc.PageHeight = 792;
            first3Doc.PageWidth = 612;
            first3Doc.MarginTop = margin;
            first3Doc.MarginBottom = margin;
            first3Doc.MarginLeft = margin;
            first3Doc.MarginRight = margin;


            Dictionary<string, List<Questions>> quotesByCompetition = new Dictionary<string, List<Questions>>();
            foreach (var q in questionsActiveList)
            {
                if (q.type == "Q")
                {
                    string competition = (tkj == "TBC") ? q.competitionTBC : q.competitionKBC;
                    if (!quotesByCompetition.ContainsKey(competition))
                        quotesByCompetition[competition] = new List<Questions>();
                    quotesByCompetition[competition].Add(q);
                }
            }

            bool firstTitle = true;
            foreach (var competition in quotesByCompetition.Keys)
            {
                if (!firstTitle)
                    first3Doc.InsertParagraph().InsertPageBreakAfterSelf();
                firstTitle = false;

                first3Doc.InsertParagraph()
                    .Append(tkj + " Quote Fill-in Competition " + competition + " - " + questionsActiveList[6].book)
                    .Font(font).FontSize(18).Bold().Alignment = Alignment.center;
                first3Doc.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize);

                foreach (var q in quotesByCompetition[competition])
                {
                    Table table1 = first3Doc.AddTable(1, 4);
                    table1.SetWidths(new float[] { 300, 300, 300, 300 });
                    table1.Rows[0].Cells[0].Paragraphs[0]
                        .Append((q.book) + " " + q.chapter + ":" + q.verse)
                        .Font(font).FontSize(18).Bold().Alignment = Alignment.center;
                    for (int i = 1; i < 4; i++)
                    {
                        table1.Rows[0].Cells[i].Paragraphs[0]
                            .Append("").Font(font).FontSize(14).Bold().Alignment = Alignment.center;
                    }
                    first3Doc.InsertTable(table1);
                    first3Doc.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize * 2);
                }

                first3Doc.InsertParagraph().InsertPageBreakAfterSelf();
                first3Doc.InsertParagraph()
                    .Append(tkj + " Quote Fill-in Competition " + competition + " - First Three Words")
                    .Font(font).FontSize(18).Bold().Alignment = Alignment.center;
                first3Doc.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize);

                foreach (var q in quotesByCompetition[competition])
                {
                    Table table2 = first3Doc.AddTable(1, 4);
                    table2.SetWidths(new float[] { 300, 300, 300, 300 });
                    table2.Rows[0].Cells[0].Paragraphs[0].Append("").Font(font).FontSize(18).Bold().Alignment = Alignment.center;
                    string[] words = q.question.Split(' ');
                    for (int i = 0; i < 3 && i < words.Length; i++)
                    {
                        table2.Rows[0].Cells[i + 1].Paragraphs[0]
                            .Append(words[i]).Font(font).FontSize(18).Bold().Alignment = Alignment.center;
                    }
                    first3Doc.InsertTable(table2);
                    first3Doc.InsertParagraph().Append("").Font(font).FontSize(spaceFontSize * 2);
                }
            }

            try
            {
                first3Doc.Save();
                System.Diagnostics.Process.Start(first3DocName);
            }
            catch (System.IO.IOException)
            {
                OpenFileException ofe = new OpenFileException();
                ofe.err(Path.GetFileName(docName));
            }
        }


        private string maskQuote(string text, int difficulty)
        {
            document.MarginTop = margin;
            document.MarginLeft = margin;
            document.MarginRight = margin;
            document.MarginBottom = margin;
            string[] words = text.Split(' ');
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < words.Length; i++)
            {
                string core = words[i].Trim();
                string punctuation = "";
                if (!string.IsNullOrEmpty(core) && char.IsPunctuation(core[core.Length - 1]))
                {
                    punctuation = core[core.Length - 1].ToString();
                    core = core.Substring(0, core.Length - 1);
                }
                bool showFirstLetter = false;
                if (difficulty == 1) showFirstLetter = true;
                if (difficulty == 2 && i % 2 == 0) showFirstLetter = true;
                if (difficulty == 3) showFirstLetter = false;
                if (!string.IsNullOrEmpty(core) && showFirstLetter)
                {
                    string masked = core[0] + new string('_', (core.Length - 1) * 3);
                    sb.Append(masked + punctuation);
                }
                else
                {
                    sb.Append(new string('_', core.Length * 3) + punctuation);
                }
                if (i < words.Length - 1)
                    sb.Append(" ");
            }
            return sb.ToString();
        }


        //private string abbreviateBookName(string bookName)
        //{
        //    Dictionary<string, string> abbreviations = new Dictionary<string, string>
        //    {
        //        { "Romans", "Rom" },
        //        { "I Timothy", "I Tim" },
        //    };

        //    if (abbreviations.ContainsKey(bookName))
        //    {
        //        return abbreviations[bookName];
        //    }
        //    return bookName;
        //}


        //<Functions for the Menu Strip>
        public void fileOpen(string textFileName)
        {
            if (textFileName.EndsWith("html"))
            {
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                System.Diagnostics.Process.Start(@"help files\" + textFileName);
            }
            else
            {
                try
                {
                    string filePath = @"data files\" + textFileName + ".txt";

                    //open it up in Excel
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    proc.EnableRaisingEvents = false;

                    proc.StartInfo.FileName = "excel.exe";
                    proc.StartInfo.Arguments = "\"" + filePath + "\"";

                    //proc.StartInfo.Verb = "open";
                    proc.Start();
                }
                catch (System.ComponentModel.Win32Exception)
                {
                    // Handle any exceptions that might arise while opening in Notepad
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();
                    //open it up in Notepad
                    proc.EnableRaisingEvents = false;
                    proc.StartInfo.FileName = @"data files\" + textFileName + ".txt";
                    proc.Start();
                }
            }
        }

        //<Functions for the Buttons>
        private void setupTSMI_Click(object sender, EventArgs e)
        {
            fileOpen("Setting Up the Program.html");
        }

        private void usingTSMI_Click(object sender, EventArgs e)
        {
            fileOpen("Using the Program.html");
        }

        private void updatingTSMI_Click(object sender, EventArgs e)
        {
            fileOpen("Updating the Files.html");
        }
        private void aboutTSMI_Click(object sender, EventArgs e)
        {
            fileOpen("About Us.html");
        }
        //Teen Flatfiles
        private void teenTSMI10_Click(object sender, EventArgs e)
        {
            fileOpen(@"Teens\Guides (25, 20, 13, 12, 10)\tbcCompetitionGuide10");
        }
        private void teenTSMI12_Click(object sender, EventArgs e)
        {
            fileOpen(@"Teens\Guides (25, 20, 13, 12, 10)\tbcCompetitionGuide12");
        }

        private void teenTSMI13_Click(object sender, EventArgs e)
        {
            fileOpen(@"Teens\Guides (25, 20, 13, 12, 10)\tbcCompetitionGuide13");
        }

        private void teenTSMI20_Click(object sender, EventArgs e)
        {
            fileOpen(@"Teens\Guides (25, 20, 13, 12, 10)\tbcCompetitionGuide20");
        }

        private void teenTSMI25_Click(object sender, EventArgs e)
        {
            fileOpen(@"Teens\Guides (25, 20, 13, 12, 10)\tbcCompetitionGuide25");
        }
        //Kid Flat Files
        private void kidTSMI20_Click(object sender, EventArgs e)
        {
            fileOpen(@"Kids\Guides (25, 20)\kbcCompetitionGuide20");
        }

        private void kidTSMI25_Click(object sender, EventArgs e)
        {
            fileOpen(@"Kids\Guides (25, 20)\kbcCompetitionGuide25");
        }
        // -------- NEW generic handler for C1‑C6 --------------
        // 1) re‑enables Submit if paths are chosen,
        // 2) refreshes the ListBox with the filtered counts.
        private void rdbCompNumber_CheckedChanged(object sender, EventArgs e)
        {
            var rb = sender as RadioButton;
            if (rb == null || !rb.Checked) return;      // ignore un‑check

            btnSubmit.Enabled = btnOutputClicked && btnInputClicked;
            countTypes();                               // update type counts
        }
        // Competition # radio‑buttons: enable Submit when paths set,
        // then refresh the type counts for the newly‑selected competition.
        private void rdbC1_CheckedChanged(object sender, EventArgs e)
        {
            if (!rdbC1.Checked) return;
            btnSubmit.Enabled = btnOutputClicked && btnInputClicked;
            countTypes();                             // update ListBox
        }
        private void rdbC2_CheckedChanged(object sender, EventArgs e)
        {
            if (!rdbC2.Checked) return;
            btnSubmit.Enabled = btnOutputClicked && btnInputClicked;
            countTypes();
        }
        private void rdbC3_CheckedChanged(object sender, EventArgs e)
        {
            if (!rdbC3.Checked) return;
            btnSubmit.Enabled = btnOutputClicked && btnInputClicked;
            countTypes();
        }
        private void rdbC4_CheckedChanged(object sender, EventArgs e)
        {
            if (!rdbC4.Checked) return;
            btnSubmit.Enabled = btnOutputClicked && btnInputClicked;
            countTypes();
        }
        private void rdbC5_CheckedChanged(object sender, EventArgs e)
        {
            if (!rdbC5.Checked) return;
            btnSubmit.Enabled = btnOutputClicked && btnInputClicked;
            countTypes();
        }
        private void rdbC6_CheckedChanged(object sender, EventArgs e)
        {
            if (!rdbC6.Checked) return;
            btnSubmit.Enabled = btnOutputClicked && btnInputClicked;
            countTypes();
        }

        private void rdb25_CheckedChanged(object sender, EventArgs e)
        {
            if (btnOutputClicked == true && btnInputClicked == true)
            {
                btnSubmit.Enabled = true;
            }
            else
            {
                return;
            }
        }

        private void rdb20_CheckedChanged(object sender, EventArgs e)
        {
            if (btnOutputClicked == true && btnInputClicked == true)
            {
                btnSubmit.Enabled = true;
            }
            else
            {
                return;
            }
        }

        private void rdb13_CheckedChanged(object sender, EventArgs e)
        {
            if (btnOutputClicked == true && btnInputClicked == true)
            {
                btnSubmit.Enabled = true;
            }
            else
            {
                return;
            }
        }

        private void rdb12_CheckedChanged(object sender, EventArgs e)
        {
            if (btnOutputClicked == true && btnInputClicked == true)
            {
                btnSubmit.Enabled = true;
            }
            else
            {
                return;
            }
        }

        private void rdb10_CheckedChanged(object sender, EventArgs e)
        {
            if (btnOutputClicked == true && btnInputClicked == true)
            {
                btnSubmit.Enabled = true;
            }
            else
            {
                return;
            }
        }



        //</Functions for the Buttons>
        //</Functions for the Menu Strip>
    }
}
