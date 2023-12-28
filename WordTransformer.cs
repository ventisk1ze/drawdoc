using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.IO;
using wp = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace DrawDocument
{
    class WordTransformer
    {
        /// <summary>
        /// Разделение документа на несколько документов, каждый не более 500 параграфов
        /// </summary>
        /// <param name="pathMainDoc">Путь, по которому лежит разделяемый файл</param>
        /// <returns>Список путей новых документов</returns>
        public static object SplitDocument(string pathMainDoc, List<string> fileParts)
        {
            try 
            { 
                int count = GetParagraphsCount(pathMainDoc);

                //Console.WriteLine($"Количество параграфов в файле {pathMainDoc} - {count}");
                //Console.WriteLine("Если количество параграфов от 500, то далее файл будет разделён на несколько частей в той же директории");
                //Console.WriteLine("Нажмите любую клавишу для продолжения");
                //Console.ReadKey();
                //Console.WriteLine();
                
                if (count >= 500)
                {
                    using (var MainDocument = WordprocessingDocument.Open(pathMainDoc, false))
                    {
                        //Делим по 500 параграфов первую половину документа
                        string TmpPath = GetTmpPath(pathMainDoc, fileParts.Count);
                        CreateDocCopy(pathMainDoc, TmpPath);
                        DeleteAfterHalf(TmpPath);
                        //количество параграфов в первой половине
                        int paragraphsCount = GetParagraphsCount(TmpPath);
                        //Console.WriteLine($"Во первой половине файла {TmpPath} количество параграфов - {paragraphsCount}");
                        if (paragraphsCount >= 500)
                        { 
                            var status = SplitDocument(TmpPath, fileParts);
                            if (status is Exception)
                                return status;
                            //удаляем промежуточный файл
                            File.Delete(TmpPath);
                        }
                        else
                            fileParts.Add(TmpPath);
                    
                        TmpPath = GetTmpPath(pathMainDoc, fileParts.Count);
                        CreateDocCopy(pathMainDoc, TmpPath);
                        DeleteBeforeHalf(TmpPath);
                        //количество параграфов в первой половине
                        paragraphsCount = GetParagraphsCount(TmpPath);
                        //Console.WriteLine($"В первой половине файла {TmpPath} количество параграфов - {paragraphsCount}");
                        if (paragraphsCount >= 500) 
                        {
                            var status = SplitDocument(TmpPath, fileParts);
                            if (status is Exception)
                                return status;
                            //удаляем промежуточный файл
                            File.Delete(TmpPath);
                        }
                        else
                            fileParts.Add(TmpPath);

                    }
                }
                else
                    fileParts.Add(pathMainDoc);

                return true;
            }
            catch (Exception ex)
            {
                return ex;
            }

            
        }

        public static void UnionDocuments(List<string> fileParts, string resultPath)
        {
            //string resultPath = ;
            File.Copy(fileParts[0], resultPath, true);
            
            using (WordprocessingDocument resultDoc =
                WordprocessingDocument.Open(resultPath, true))
            {
                for (int i = fileParts.Count - 1; i > 0; --i)
                {
                    string altChunkId = "AltChunkId" + i;
                    MainDocumentPart mainPart = resultDoc.MainDocumentPart;
                    AlternativeFormatImportPart chunk =
                        mainPart.AddAlternativeFormatImportPart(
                        AlternativeFormatImportPartType.WordprocessingML, altChunkId);

                    string filename = fileParts[i];

                    using (FileStream fileStream = File.Open(filename, FileMode.Open))
                        chunk.FeedData(fileStream);
                    wp.AltChunk altChunk = new wp.AltChunk();
                    altChunk.Id = altChunkId;
                    mainPart.Document
                        .Body
                        .InsertAfter(altChunk, mainPart.Document.Body
                        .Elements<wp.Paragraph>().Last());
                    mainPart.Document.Save();                  
                }

                resultDoc.Save();
                DeleteFiles(fileParts);
            }
        }

        /// <summary>
        /// Удаляет все файлы из списка путей
        /// </summary>
        /// <param name="files"></param>
        private static void DeleteFiles(List<string> files)
        {
  
            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);

                if (fi.Exists)
                    File.Delete(file);
            }
        }

        private static void DeleteAfterHalf(string path)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(path, true))
            {
                wp.Document doc = document.MainDocumentPart.Document;

                List<wp.Text> textparts = document.MainDocumentPart.Document.Body.Descendants<wp.Text>().ToList();
                int halfTextpartNumber = textparts.Count / 2;

                for (int i = halfTextpartNumber; i < textparts.Count; i ++)
                { 
                    wp.Text textfield = textparts[i];
                    RemoveItem(textfield);
                }
                
                document.Save();

            }
        }

        /// <summary>
        /// Удаление элемента документа
        /// </summary>
        /// <param name="path">Путь к документу</param>
        private static void DeleteBeforeHalf(string path)
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(path, true))
            {
                wp.Document doc = document.MainDocumentPart.Document;

                List<wp.Text> textparts = document.MainDocumentPart.Document.Body.Descendants<wp.Text>().ToList();
                int halfTextpartNumber = textparts.Count / 2;
                halfTextpartNumber = DecreaseWhileNewParent(textparts, halfTextpartNumber);
                if (halfTextpartNumber == 0)
                    throw new Exception("Похоже, алгоритм не пригоден для разрезки таких файлов");
                    //Console.WriteLine("Похоже, алгоритм не пригоден для разрезки таких файлов");
                for (int i = 0; i <= halfTextpartNumber; ++ i)
                { 
                    wp.Text textfield = textparts[i];
                    RemoveItem(textfield);
                 }
                
                document.Save();
                
            }
        }

        /// <summary>
        /// Так как у halfTextNumber-1 может оказаться тот же прародитель, то мы удалим уже удалённое в первой половине. Попробуем этого не допустить
        /// </summary>
        /// <param name="textparts"></param>
        /// <param name="halfTextpartNumber"></param>
        /// <returns></returns>
        private static int DecreaseWhileNewParent(List<wp.Text> textparts, int halfTextpartNumber)
        {
            //получим прародителя центрального элемента
            wp.Text textfield = textparts[halfTextpartNumber];
            DocumentFormat.OpenXml.OpenXmlElement parent_textfield = GetBodyParentElement(textfield);

            int i = halfTextpartNumber;
            //пока не найдём другого родителя, ищем
            bool find = false;
            while (i > 0 && !find)
            {
                i--;
                wp.Text textfield_before = textparts[i];
                DocumentFormat.OpenXml.OpenXmlElement parent_textfield_before = GetBodyParentElement(textfield_before);
                if (!parent_textfield.Equals(parent_textfield_before))
                    find = true;
                
            }

            return i;
        }

        private static OpenXmlElement GetBodyParentElement(wp.Text item)
        {
            DocumentFormat.OpenXml.OpenXmlElement element = item;

            while (!(element.Parent is wp.Body) && element.Parent != null)
            {
                element = element.Parent;
            }

            return element;
        }

        /// <summary>
        /// Удаление элемента документа со всеми родителями
        /// </summary>
        /// <param name="path">Элемент документа</param>
        private static void RemoveItem(wp.Text item)
        {
            DocumentFormat.OpenXml.OpenXmlElement element = item;
            while (!(element.Parent is wp.Body) && element.Parent != null)
            {
                element = element.Parent;
            }

            if (element.Parent != null)
            {
                element.Remove();
            }
        }

        /// <summary>
        /// Get path for temp file = filename + '_tmp' + number
        /// </summary>
        /// <param name="filename">filename for build temp filename</param>
        /// <param name="number">number for build temp filename</param>
        /// <returns></returns>
        private static string GetTmpPath(string filename, int number)
        {
            return Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + "_tmp" + number + Path.GetExtension(filename));
        }

        /// <summary>
        /// Get count of paragraphs by document
        /// </summary>
        /// <param name="path">path to doc</param>
        /// <returns>Count of paragraphs</returns>
        public static int GetParagraphsCount(string path)
        {
            int count = 0;
            using (var doc = WordprocessingDocument.Open(path, false))
            {
                var paragraphs = doc.MainDocumentPart.Document.Descendants<wp.Paragraph>();
                
                foreach (var paragraph in paragraphs)
                {
                    count++;
                }
            }

            return count;
        }

        /// <summary>
        /// Создаёт копию файла
        /// </summary>
        /// <param name="inputPath">Путь к файлу</param>
        /// <param name="outputPath">Путь к копии</param>
        private static void CreateDocCopy(string inputPath, string outputPath)
        {
            File.Copy(inputPath, outputPath, true);

            //using (var mainDoc = WordprocessingDocument.Open(inputPath, false))
            //using (var resultDoc = WordprocessingDocument.Create(outputPath,
            //  DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            //{
            //    copy parts from source document to new document
            //    foreach (var part in mainDoc.Parts)
            //        resultDoc.AddPart(part.OpenXmlPart, part.RelationshipId);
                //resultDoc.Save();
            //}
        }


    }
}
