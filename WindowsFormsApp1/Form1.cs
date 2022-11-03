using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.Speech.Recognition;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Diagnostics;
using System.Windows.Forms.VisualStyles;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        SpeechSynthesizer ss = new SpeechSynthesizer();
        static readonly CultureInfo _language = new CultureInfo("ru-RU");
        DirectoryInfo directoryInfoFolder;
        List<string> nameForDirList = new List<string>(); 
        List<string> nameForFileList = new List<string>();
        SpeechRecognitionEngine sre = new SpeechRecognitionEngine(_language);
        Process process = new Process();


        public Form1()
        {
            InitializeComponent();
            
            CheckForIllegalCrossThreadCalls = false;

            new Thread(() =>
            {
                Action action = () =>
                {
                    ss.SetOutputToDefaultAudioDevice();
                    ss.Volume = 100;// от 0 до 100 громкость голоса
                    ss.Rate = 2; //от -10 до 10 скорость голоса

                    sre.SetInputToDefaultAudioDevice();
                    sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Sre_SpeechRecognized);

                    sre.LoadGrammar(PersonalGrammar.ChoiseGrammar());
                    sre.LoadGrammar(PersonalGrammar.StandartPathGrammar());
                    sre.LoadGrammar(PersonalGrammar.GoToBackGrammar());

                    sre.RecognizeAsync(RecognizeMode.Multiple);
                };
                if (InvokeRequired)
                    Invoke(action);
                else
                    action();

            }).Start();
        }

        private void Sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string _SpokenText = e.Result.Text; //сказанный текст
            float confidence = e.Result.Confidence; //Точность сказаного текста(процент совпадения)

            if (confidence >= 0.60)
            {
                if (_SpokenText.IndexOf($"выбрать папку") >= 0)
                {
                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            ChoiseDir();
                        };
                        Invoke(action);


                    }).Start();
                }
                if (_SpokenText.IndexOf($"открыть папку") >= 0)
                {
                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            OpenDir(_SpokenText);
                        };
                        Invoke(action);


                    }).Start();
                }
                if (_SpokenText.IndexOf($"открыть файл") >= 0 || _SpokenText.IndexOf($"закрыть файл") >= 0)
                {
                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            OpenFile(_SpokenText);
                        };
                        Invoke(action);
                    }).Start();
                }
                if (_SpokenText.IndexOf($"открыть стартовую папку") >= 0)
                {
                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            StandartPath();
                        };
                        Invoke(action);
                    }).Start();
                }
                if (_SpokenText.IndexOf($"выйти из папки") >= 0)
                {
                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            GoToBack();
                        };
                        Invoke(action);
                    }).Start();
                }

                if (_SpokenText.IndexOf($"изменить файл") >= 0)
                {
                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            RenameFiles(_SpokenText);
                        };
                        Invoke(action);
                    }).Start();
                }
                if (_SpokenText.IndexOf($"переместить файл") >= 0)
                {
                    new Thread(() =>
                    {
                        Action action = () =>
                        {
                            MoveFile(_SpokenText);
                        };
                        Invoke(action);
                    }).Start();
                }

            }
        }
        public void RenameFiles(string _nameFile)
        {
            textBox2.Text = _nameFile;
            try
            {
                string[] arrayText = _nameFile.Split(' ');
                string fullPath = directoryInfoFolder + $"\\{arrayText[2]}";

               
                string text = arrayText[2];
                string textEnd = arrayText[2];
                int dotIndex = text.IndexOf('.');
                if (dotIndex >= 0)
                {
                    text = text.Substring(0, dotIndex);
                    textEnd = textEnd.Substring(dotIndex);
                }
                string txt = directoryInfoFolder +"\\\\"+ text + "_this_file_needs_to_be_changed" + textEnd;
                File.Move(fullPath, txt);  
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Неудалось пометить {_nameFile} на редактирование, Ошибка: " + ex.Message);
            }
        }

        public void StandartPath()       
        {
            string[] Drives = Environment.GetLogicalDrives();

            string urlParent = Drives[1];
            webBrowser.Url = new Uri(urlParent);
            textBox1.Text = urlParent;
            directoryInfoFolder = new DirectoryInfo(urlParent);

            CreateLocalGrammar();
        }

        public void OpenFile(string _nameFile)
        {
            textBox2.Text = _nameFile;
            
            try
            {
                string[] arrayText = _nameFile.Split(' ');

                if (arrayText[0] == "открыть")
                {
                    process.StartInfo.FileName = directoryInfoFolder.FullName + $"\\{arrayText[2]}";
                    process.Start();
                }
                if (arrayText[0] == "закрыть")
                {
                    Process[] runningProcesses = Process.GetProcesses();
                    foreach (Process process in runningProcesses)
                    {     
                        foreach (ProcessModule module in process.Modules)
                        {
                            if (module.FileName.Equals($"{arrayText}.png"))
                            {
                                process.Kill();
                            }
                        }
                    }
                }   
            }
            catch (Exception ex)
            {
                MessageBox.Show("Файл с таким названием не существует((" + ex.Message);
            }
        }

        public void GoToBack()
        {
            try
            {
                if (directoryInfoFolder.Root.ToString() == directoryInfoFolder.ToString())
                {
                    MessageBox.Show("Выйти из диска невозможно/ помейте диск командой \"Поменять диск\"");
                    return;
                }
                string urlParent = Directory.GetParent(directoryInfoFolder.ToString()).FullName;
                webBrowser.Url = new Uri(urlParent);
                textBox1.Text = urlParent;
                directoryInfoFolder = new DirectoryInfo(urlParent);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Выйти из это папки по каким то причинам невозможно: " + ex.Message);
            }
        }

        public void MoveFile(string _nameFile)
        {
            try
            {
                string[] arrayText = _nameFile.Split(' ');
                string path = directoryInfoFolder.FullName + $"\\{arrayText[2]}";

                string newPath;
                new Thread(() =>
                {
                    Action action = () =>
                    {
                        using (FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Select your path." })
                        {
                            if (fbd.ShowDialog() == DialogResult.OK)
                            {
                                newPath = fbd.SelectedPath;
                                FileInfo fileInf = new FileInfo(path);
                                if (fileInf.Exists)
                                {
                                    File.Move(path, newPath+ $"\\{arrayText[2]}");
                                }
                            }
                        }
                    };
                    if (InvokeRequired)
                        Invoke(action);
                    else
                        action();
                }).Start();
                
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Переместить файл по каким то причинам невозможно: " + ex.Message);
            }
        }

        public void OpenDir(string _nameDir)
        {
            textBox2.Text = _nameDir;
            try
            {               
                string[] arrayText = _nameDir.Split(' ');
               
                string url;
                if (directoryInfoFolder.Root.ToString() == directoryInfoFolder.ToString())
                {
                    url = directoryInfoFolder.FullName + $"{arrayText[2]}";
                }
                else
                {
                    url = directoryInfoFolder.FullName + $"\\{arrayText[2]}";
                }
                webBrowser.Url = new Uri(url);
                textBox1.Text = url;
                directoryInfoFolder = new DirectoryInfo(url);

                CreateLocalGrammar();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Файл не существует или скрыт " + ex.Message);
            }
        }

        public void ChoiseDir()
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Select your path." })
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {

                    webBrowser.Url = new Uri(fbd.SelectedPath);
                    directoryInfoFolder = new DirectoryInfo(fbd.SelectedPath);
                    textBox1.Text = fbd.SelectedPath;
                }
            }
            CreateLocalGrammar();
        }

        public void CreateLocalGrammar()
        {
            try
            {
                //Создаём грамматику со всеми названиями папок в выбранной папке
                foreach (DirectoryInfo direct in directoryInfoFolder.GetDirectories())
                {
                    nameForDirList.Add(direct.Name);
                }

                GrammarBuilder grammarBuilder = new GrammarBuilder(); 
                grammarBuilder.Culture = _language;
                grammarBuilder.Append("открыть"); 
                grammarBuilder.Append("папку"); 
                grammarBuilder.Append(new Choices(nameForDirList.ToArray()));
                sre.LoadGrammar(new Grammar(grammarBuilder));


                //Создаём грамматику со всеми названиями файлов в выбранной папке
                foreach (FileInfo _file in directoryInfoFolder.GetFiles())
                {
                    nameForFileList.Add(_file.Name);
                }

                GrammarBuilder grammarBuilder1 = new GrammarBuilder(); 
                grammarBuilder1.Culture = _language;
                grammarBuilder1.Append(new Choices("открыть", "закрыть"));
                grammarBuilder1.Append("файл");
                grammarBuilder1.Append(new Choices(nameForFileList.ToArray()));
                sre.LoadGrammar(new Grammar(grammarBuilder1));

                GrammarBuilder grammarBuilder2 = new GrammarBuilder();
                grammarBuilder2.Culture = _language;
                grammarBuilder2.Append(new Choices("изменить"));
                grammarBuilder2.Append("файл");
                grammarBuilder2.Append(new Choices(nameForFileList.ToArray()));
                sre.LoadGrammar(new Grammar(grammarBuilder2));

                GrammarBuilder grammarBuilder3 = new GrammarBuilder();
                grammarBuilder3.Culture = _language;
                grammarBuilder3.Append(new Choices("переместить"));
                grammarBuilder3.Append("файл");
                grammarBuilder3.Append(new Choices(nameForFileList.ToArray()));
                sre.LoadGrammar(new Grammar(grammarBuilder3));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }


        void button1_Click(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            new Thread(()=>
            {
                Action action = () =>
                {
                    StandartPath();
                };
                if (InvokeRequired)
                    Invoke(action);
                else
                    action();
            }).Start();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            new Thread(() =>
            {
                Action action = () =>
                {
                    OpenDir("Открыть папку Users");
                };
                if (InvokeRequired)
                    Invoke(action);
                else
                    action();
            }).Start();
            
        }
        private void button3_Click(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            new Thread(() =>
            {
                Action action = () =>
                {
                    GoToBack();
                };
                if (InvokeRequired)
                    Invoke(action);
                else
                    action();
            }).Start();   
        }
        private void button4_Click(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            new Thread(() =>
            {
                Action action = () =>
                {
                    RenameFiles("\\sp5.dock");
                };
                
                action();
            }).Start();
        }
    }
}
