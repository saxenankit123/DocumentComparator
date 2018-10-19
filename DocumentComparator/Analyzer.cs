using log4net;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DocumentComparator
{
    public class Analyzer
    {
        private ILog log = LogManager.GetLogger("Analyzer");
        private Dictionary<string, string> originalDocDictionary;
        private Dictionary<string, string> newDocDictionary;

        public Analysis doAnalysis(string originalFileName,string newFileName)
        {
            log.Info("Starting Document Comparision");
            log.Debug("Original File - " + originalFileName);
            log.Debug("New File - " + newFileName);

            Document wDocNew = null;
            Document wDocOriginal = null;
            Analysis result = null;
            
            try
            {
                log.Info("Loading New File");
                wDocNew = new Microsoft.Office.Interop.Word.Application().Documents.Open(newFileName, ReadOnly: true);
                log.Info("Loading Original File");
                wDocOriginal = new Microsoft.Office.Interop.Word.Application().Documents.Open(originalFileName, ReadOnly: true);
                log.Info("Preparing New File");
                prepareNewDoc(wDocNew);
                log.Info("Preparing Original File");
                prepareOriginalDoc(wDocOriginal);

                result = new Analysis();
                log.Info("Analzing..");
                foreach (KeyValuePair<string, string> new_item in newDocDictionary)
                {

                    if (!originalDocDictionary.ContainsKey(new_item.Key)) // New clause added
                    {
                        result.newClauses.Add(new_item.Key.ToString(), new_item.Value);
                    }
                    else if (!new_item.Value.ToString().Equals(originalDocDictionary[new_item.Key]))
                    {
                        //clause number is present but content is not equal
                        if (new_item.Value.ToString().Trim().Equals(""))
                        {
                            //clause is deleted
                            result.deletedClauses.Add(new_item.Key.ToString(), originalDocDictionary[new_item.Key]);
                        }
                        else
                        {
                            //clause is modified
                            result.modifiedClauses.Add(new_item.Key.ToString(), new_item.Value);
                        }


                    }
                }
            }
            catch (Exception exception)
            {
                log.Error("An error occured while processing - " + exception.Message);
                log.Error(exception);
                return null;
            }
            finally
            {
                log.Info("Analysis Complete");
                wDocNew.Close();
                wDocOriginal.Close();
            }

            return result;

        }

        private void prepareOriginalDoc(Document wDocOriginal)
        {
            originalDocDictionary = new Dictionary<string, string>();
            foreach (Paragraph p in wDocOriginal.Paragraphs)
            {
                originalDocDictionary.Add(p.Range.ListFormat.ListString, p.Range.Text); //if the clauses are not numbered then we can get issue...needs to be handled

            }

        }

        private void prepareNewDoc(Document wDocNew)
        {
            newDocDictionary = new Dictionary<string, string>();
            foreach (Paragraph p in wDocNew.Paragraphs)
            {
                newDocDictionary.Add(p.Range.ListFormat.ListString, p.Range.Text); //if the clauses are not numbered then we can get issue...needs to be handled

            }
        }
    }
}
