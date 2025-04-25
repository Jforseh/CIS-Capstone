using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;

namespace BibleCompiler2
{
    internal class Questions
    {
        public string competitionTBC, competitionKBC, competitionJMP, book, chapter, verse, type, question, answer;
        public Questions(string competitionTBC, string competitionKBC, string competitionJMP, string book, string chapter, string verse, string type, string question, string answer)
        {
            if(type == "M")
            {
                int num = 0;
                string tempAns = "";
                string tempString = answer;
                for (int k = 0; k < tempString.Length; k++)
                {
                    if (int.TryParse(tempString[k].ToString(), out num)&& num > 1)
                    {
                        tempAns += '\n';
                    }
                    tempAns += tempString[k];
                }
                answer = tempAns;
            }
            this.competitionTBC = competitionTBC;
            this.competitionKBC = competitionKBC;
            this.competitionJMP = competitionJMP;
            this.book = book;
            this.chapter = chapter;
            this.verse = verse;
            this.type = type;
            this.question = question;
            this.answer = answer;

            


        }
        public override string ToString() 
        {
            return(book + " " + chapter + ":" + verse + " " + type + " " + question + " / " + answer).ToString();
        }

    }
}
