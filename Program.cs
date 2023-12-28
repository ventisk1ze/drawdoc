using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using System.Text;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json;

using System;

namespace DrawDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            Logger.TextFile logger = new Logger.TextFile(Path.Combine("paintlog", String.Format("DrawDocument_log_{0}.txt", DateTime.Now.Date.ToString("yyyy-MM-dd"))));
            logger.Add("Программа начинает работать!");

            string json_path;//= @"C:\Users\Xiaomi\source\repos\test.json";
            if (args.Length != 1)
            {
                logger.Add("Неверное число аргументов на вход программы!");
                logger.Add("Программа закончила работать.");
                return;
            }
            else
            {
                json_path = args[0];
                if (Path.GetExtension(json_path) != ".json")
                {
                    logger.Add("Расширение файла-аргумента должно быть .json!");
                    logger.Add("Программа закончила работать.");
                    return;
                }
            }

            List<DrawParams> files;
            try
            {
                string json_string = File.ReadAllText(json_path, Encoding.UTF8);
                files = JsonConvert.DeserializeObject<List<DrawParams>>(json_string);
            }
            catch (Exception ex)
            {
                logger.Add(String.Format("Ошибка при десериализации json-файла: {0}", ex.ToString()));
                logger.Add("Программа закончила работать.");
                return;
            }
            try
            {
                foreach (var file in files)
                {
                    List<string> fileParts = new List<string>();

                    int count = WordTransformer.GetParagraphsCount(file.InputPath);
                    logger.Add(String.Format("В файле {0} - {1} параграфов", file.InputPath, count));

                    object status = WordTransformer.SplitDocument(file.InputPath, fileParts);

                    if (status is Exception)
                    {
                        logger.Add(String.Format("Разделить файл {0} на части не удалось: {1}", file.InputPath, status.ToString()));
                        fileParts.Clear();
                        fileParts.Add(file.InputPath);
                    }
                    else
                    {
                        logger.Add(String.Format("Файл {0} разделён на {1} части", file.InputPath, fileParts.Count));
                    }

                    foreach (string filePart in fileParts)
                    { 
                        Document document = new Document();
                        document.LoadFromFile(filePart);
                        foreach (Sentence sentence in file.Sentences)
                        {
                            foreach (string txt in sentence.Text.Split("\n"))
                            {
                                TextSelection[] text = document.FindAllString(txt, false, false);
                                if (text != null)
                                {
                                    foreach (TextSelection selection in text)
                                    {
                                        selection.GetAsOneRange().CharacterFormat.TextColor = Color.FromName(sentence.Color);
                                    }
                                }
                                else
                                    logger.Add(String.Format(@"Не удалось найти в файле {0} cтроку из json:" + Environment.NewLine + "'{1}'", filePart, txt));

                            }
                        }
                        if (fileParts.Count == 1)
                            document.SaveToFile(file.OutputPath, FileFormat.Docx);
                        else
                            document.SaveToFile(filePart);
                    }

                    if (fileParts.Count > 1)
                    {
                        WordTransformer.UnionDocuments(fileParts, file.OutputPath);
                    }
                }
            }

            catch (Exception ex)
            {
                logger.Add(String.Format("Ошибка при попытке раскраски файла: {0}", ex.ToString()));
            }

            logger.Add("Программа закончила работать.");
        }
    }
}
