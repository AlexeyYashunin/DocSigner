using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.IO;
using System;

using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

using iTextSharp.text.pdf;
using Aspose.Pdf.Text;

using MoreLinq;

namespace DocSigner {
    public partial class DocSigner_Window : Form {
        //список pdf файлов в указанной папке
        public List<string> files_spisok = null;//входные документы

        public int done_files_amount = 0;//кол-во сделанных файлов
        public int done_files_percent = 0;//процент сделанных файлов

        public int curIndFileNames = 0;//текущий индекс в списке файлов, для лога

        public string currentPath = System.IO.Directory.GetCurrentDirectory();

        public Task[] tasks_prostavit_podpisi_i_pechati_aspose = null;

        public int delimTaskValue = 1;//кол-во тасков запускаемых одновременно

        static object locker_task_get_podpisanti_data_from_pdf_file = new object();
        static object locker_prostavit_podpisi_i_pechati_itext_seven_task = new object();
        static object locker_prostavit_pechat_task = new object();

        //сохраняем путь к папке входных данных
        public string input_document_file = "";

        //путь к файлу со списком документов для подписания
        public string vhodnie_dannie_papka = System.IO.Directory.GetCurrentDirectory() + @"\Source\Documenti";
        public string pechati_papka = System.IO.Directory.GetCurrentDirectory() + @"\Source\Pechati\Image";
        public string podpisanti_papka = System.IO.Directory.GetCurrentDirectory() + @"\Source\Podpisi";
        public string pechat_organizatcii_file_path = System.IO.Directory.GetCurrentDirectory() + @"\Source\Pechati\Pechat_organizatcii.png";
        public string pechat_dlya_protocolov_file_path = System.IO.Directory.GetCurrentDirectory() + @"\Source\Pechati\Pechat_dlya_protocolov.png";
        public string json_papka = System.IO.Directory.GetCurrentDirectory() + @"\Source\JSON";

        public string outCrashFolder = System.IO.Directory.GetCurrentDirectory() + "\\Out" + "\\Log";
        public string outCrashFile = System.IO.Directory.GetCurrentDirectory() + "\\Out" + "\\Log" + "\\" + "Crash_log.txt";

        public string outFilePath = System.IO.Directory.GetCurrentDirectory() + "\\Out";
        public string out_img_podpisi_FilePath = System.IO.Directory.GetCurrentDirectory() + "\\Out\\[IMG]Podpisi_temp";

        public DocSigner_Window() {
            InitializeComponent();

            LicenseHelper.ModifyInMemory.ActivateMemoryPatching();

            //настраиваем worker'а
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
        }

        private void OcrBtn_Click(object sender, EventArgs e) {
            if(backgroundWorker1.IsBusy) {
                return;
            }

            //берем папку с pdfками
            if(folderBrowserDialog.ShowDialog() == DialogResult.OK) {
                vhodnie_dannie_papka = folderBrowserDialog.SelectedPath;//берем файлик со списком доков для распознания
            }

            progress_lbl.Text = "0 %";
            process_pb.Value = 0;

            done_files_amount = 0;

            //запускаем worker'а
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e) {
           if((done_files_percent == 0 && e.ProgressPercentage != 0) ||
               (done_files_percent < e.ProgressPercentage)
            ) {
                done_files_percent = e.ProgressPercentage;
            }

            process_pb.Value = done_files_percent;
            progress_lbl.Text = done_files_percent.ToString() + " %";
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e) {
            if(folderBrowserDialog.SelectedPath == "") {
                return;
            }

            //берем кол-во тасков   
            if(tasks_tb.Text != "") {
                try {
                    delimTaskValue = Convert.ToInt32(tasks_tb.Text);

                    if(delimTaskValue <= 0) {
                        delimTaskValue = 1;

                        MessageBox.Show("Only numbers > 0 \"tasks\"", "Wrong input");
                    }

                } catch(System.Exception ex) {
                    delimTaskValue = 1;

                    MessageBox.Show("Only numbers in \"tasks\"", "Wrong input");

                    return;
                }
            }

            create_folder(outFilePath);//создаем папку для выходных документов pdf
            create_folder(outCrashFolder);//папка с крашами

            //проходим по страницам документа
            List<S_Pdf_Page_Info> infoPages = new List<S_Pdf_Page_Info>();//информация по страницам pdf
            List<S_Podpisant> podpisanti = get_podpisanti_data(podpisanti_papka);//подписанты


            //обходим список файлов
            List<string> files_spisok = getPdfFiles(vhodnie_dannie_papka);

            tasks_prostavit_podpisi_i_pechati_aspose = new Task[files_spisok.Count];

            int progress_percent = 0;

            for(int i = 0; i < files_spisok.Count; i++) {
                object index_page = i;

                tasks_prostavit_podpisi_i_pechati_aspose[i] = new Task(() => task_prostavit_podpisi_i_pechati_aspose(files_spisok[(int)index_page], ref podpisanti, index_page, files_spisok.Count));
            }

            run_tasks(tasks_prostavit_podpisi_i_pechati_aspose, delimTaskValue, files_spisok.Count);//запускаем таски

            //ждем завершения всех тасков
            Task.WaitAll(tasks_prostavit_podpisi_i_pechati_aspose);

            progress_percent = 100;

            backgroundWorker1.ReportProgress(progress_percent);
        }

        public void run_tasks(Task[] tasks, int delimTaskValue, int porogovoe_znachenie_zapyska) {
            //запускаем таски
            for(int j = 0; j < tasks.Length; j++) {
                if(((j + 1) % delimTaskValue == 0)) {
                    Task[] delimTasks = new Task[delimTaskValue];

                    for(int k = 0; k < delimTaskValue; k++) {
                        delimTasks[k] = tasks[(j + 1) - delimTaskValue + k];
                    }

                    for(int k = 0; k < delimTaskValue; k++) {
                        delimTasks[k].Start();
                    }

                    //ждем завершения всех тасков
                    Task.WaitAll(delimTasks);
                } else if(delimTaskValue > porogovoe_znachenie_zapyska) {//иначе запускаем все таски
                    Task[] delimTasks = new Task[porogovoe_znachenie_zapyska];

                    for(int k = 0; k < porogovoe_znachenie_zapyska; k++) {
                        delimTasks[k] = tasks[k];
                    }

                    for(int k = 0; k < porogovoe_znachenie_zapyska; k++) {
                        delimTasks[k].Start();
                    }

                    //ждем завершения всех тасков
                    Task.WaitAll(delimTasks);

                    break;
                }
            }

            if((delimTaskValue > porogovoe_znachenie_zapyska) && (tasks.Length > 2)) {
                //разбираем остаточную страницу(если есть)
                int ostatokPages = (tasks.Length % 2);

                if(ostatokPages != 0) {
                    try{
                        tasks[tasks.Length - 1].Start();
                        tasks[tasks.Length - 1].Wait();
                    } catch (System.Exception ex){                    	
                    }                    
                }
            }
        }

        public void task_prostavit_podpisi_i_pechati_aspose(string input_file_path, ref List<S_Podpisant> podpisanti, object index_page, int porogovoe_znachenie_zapyska) {
            bool is_podpis_i_pechati_prostavleni = false;

            int progress_percent = 0;//текущее значение обработки в процентах

            is_podpis_i_pechati_prostavleni = prostavit_podpisi_i_pechati_aspose(input_file_path, ref podpisanti);//проставляем подписи и печати

            result_print(input_file_path, is_podpis_i_pechati_prostavleni);

            done_files_amount += 1;

            progress_percent = ((done_files_amount * 100) / porogovoe_znachenie_zapyska);

            backgroundWorker1.ReportProgress(progress_percent);
        }

        public bool prostavit_podpisi_i_pechati_aspose(string input_pdf_file_path, ref List<S_Podpisant> podpisanti) {
            try {
                S_Pdf_Document pdf_Document = new S_Pdf_Document();

                pdf_Document.input_pdf_file_path = input_pdf_file_path;

                string imia_file = Path.GetFileName(input_pdf_file_path);//имя входного файла

                //создаем выходную директорию
                string folder_name = Path.GetFileName(Path.GetDirectoryName(vhodnie_dannie_papka + "\\"));
                int index_of_vhodnaya_papka = input_pdf_file_path.IndexOf(folder_name);
                string ierarhia_papok = Path.GetDirectoryName(input_pdf_file_path.Substring(index_of_vhodnaya_papka));

                string out_folder_path = outFilePath + "\\" + ierarhia_papok;

                create_folder(out_folder_path);

                string out_pdf_file = out_folder_path + "\\" + imia_file;

                //если файл уже делали, то пропускаем
                if(File.Exists(out_pdf_file) == true) {
                    return true;
                }

                pdf_Document.pages = new List<S_Pdf_Page_Info>();

                List<S_Pdf_Page_Info> pages = new List<S_Pdf_Page_Info>();

                Aspose.Pdf.Document input_pdf_aspose_doc = new Aspose.Pdf.Document(input_pdf_file_path);

                //сопоставляем подписанта и подпись
                for(int i = 0; i < input_pdf_aspose_doc.Pages.Count; i++) {
                    S_Pdf_Page_Info pdf_Page_Info = new S_Pdf_Page_Info();

                    pdf_Page_Info.input_pdf_file_path = input_pdf_file_path;
                    pdf_Page_Info.numberOfPDFPage_ = i + 1;

                    List<S_Podpis_on_page_data> podpisi = new List<S_Podpis_on_page_data>();

                    //ищем подписантов
                    for(int j = 0; j < podpisanti.Count; j++) {
                        List<S_Podpis_on_page_data> temp_podpisi = new List<S_Podpis_on_page_data>();

                        temp_podpisi = get_podpis_on_page_data(input_pdf_aspose_doc, input_pdf_file_path, podpisanti[j], i + 1);

                        if(temp_podpisi.Count > 0) {
                            for(int s = 0; s < temp_podpisi.Count; s++) {
                                podpisi.Add(temp_podpisi[s]);
                            }
                        }
                    }

                    //ищем средний размер изображения подписи по имеющимся данным
                    double average_delta_x = 0;
                    double average_delta_y = 0;

                    if(podpisi.Count > 0) {
                        try {
                            average_delta_x = podpisi.Where(podpis => podpis.podpis_coordinates != null).Where(podpis_rect => podpis_rect.podpis_coordinates.Rectangle != null).Average(p => p.podpis_coordinates.Rectangle.Width);
                        } catch(System.Exception ex) {
                            average_delta_x = podpisi[0].podpisant_coordinates.Rectangle.Width;
                        }

                        average_delta_y = podpisi.Where(podpis => podpis.podpisant_coordinates != null).Where(podpis_rect => podpis_rect.podpisant_coordinates.Rectangle != null).Average(p => p.podpisant_coordinates.Rectangle.Height);

                        double podpis_image_width = average_delta_x;
                        podpis_image_width *= 3.5;

                        double podpis_image_height = average_delta_y;
                        podpis_image_height *= 3.5;

                        //устанавливаем координаты изображения
                        for(int r = 0; r < podpisi.Count; r++) {
                            Aspose.Pdf.Rectangle podpis_image_coordinates = null;

                            try {
                                podpis_image_coordinates = new Aspose.Pdf.Rectangle(
                                    podpisi[r].podpis_coordinates.Rectangle.LLX + (podpisi[r].podpis_coordinates.Rectangle.Width / 2) - (podpis_image_width / 2),
                                    podpisi[r].podpisant_coordinates.Rectangle.LLY,
                                    podpisi[r].podpis_coordinates.Rectangle.LLX + (podpisi[r].podpis_coordinates.Rectangle.Width / 2) - (podpis_image_width / 2) + podpis_image_width,
                                    podpisi[r].podpisant_coordinates.Rectangle.LLY + podpis_image_height
                                );
                            } catch(System.Exception ex) {
                                if(i == 0) {//если 1я страница
                                    podpis_image_coordinates = new Aspose.Pdf.Rectangle(
                                        podpisi[r].podpisant_coordinates.Rectangle.LLX - (podpisi[r].podpisant_coordinates.Rectangle.Width * 2),
                                        podpisi[r].podpisant_coordinates.Rectangle.LLY,
                                        podpisi[r].podpisant_coordinates.Rectangle.LLX - (podpisi[r].podpisant_coordinates.Rectangle.Width * 2) + podpis_image_width,
                                        podpisi[r].podpisant_coordinates.Rectangle.LLY + podpis_image_height
                                    );
                                } else {
                                    podpis_image_coordinates = new Aspose.Pdf.Rectangle(
                                        podpisi[r].podpisant_coordinates.Rectangle.LLX - (podpisi[r].podpisant_coordinates.Rectangle.Width * 3.5),
                                        podpisi[r].podpisant_coordinates.Rectangle.LLY,
                                        podpisi[r].podpisant_coordinates.Rectangle.LLX - (podpisi[r].podpisant_coordinates.Rectangle.Width * 3.5) + podpis_image_width,
                                        podpisi[r].podpisant_coordinates.Rectangle.LLY + podpis_image_height
                                    );
                                }
                            }

                            podpisi[r].podpis_image_coordinates = podpis_image_coordinates;
                        }

                        if(podpisi.Count > 0) {
                            if(i > 1) {//с 3й странцы проверяем
                                if(((get_podpisi_count(input_pdf_aspose_doc, i) > pdf_Document.pages[i - 1].podpisi.Count) ||
                                    (pdf_Document.pages[i - 1].podpisi.Count == 1)) && (podpisi[0].podpis_coordinates != null)
                                ) {
                                    try {
                                        pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_coordinates = (TextFragment)podpisi[0].podpis_coordinates.Clone();//последней подписи на предыдущей странице ставим координаты первой подписи с текущей
                                        pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_image_coordinates = (Aspose.Pdf.Rectangle)podpisi[0].podpis_image_coordinates.Clone();

                                        pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_image_coordinates.LLX = podpisi[0].podpis_coordinates.Rectangle.LLX + (podpisi[0].podpis_coordinates.Rectangle.Width / 2) - (podpis_image_width / 2);
                                        pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_image_coordinates.LLY = pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpisant_coordinates.Rectangle.LLY;
                                        pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_image_coordinates.URX = podpisi[0].podpis_coordinates.Rectangle.URX - (podpisi[0].podpis_coordinates.Rectangle.Width / 2) - (podpis_image_width / 2) + podpis_image_width;
                                        pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_image_coordinates.URY = pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpisant_coordinates.Rectangle.LLY + podpis_image_height;
                                    } catch (System.Exception ex){                                    	
                                    }                                   
                                }
                            }
                        }
                    }

                    if(podpisi.Count == 0) {//если подписантов на текущей странице нету, а на предыдущей странице слова подпись нет
                        if(i > 1) {//с 3й странцы проверяем
                            if(((get_podpisi_count(input_pdf_aspose_doc, i) > pdf_Document.pages[i - 1].podpisi.Count) ||
                                (pdf_Document.pages[i - 1].podpisi.Count == 1))
                            ) {//если самих слов "подпись" меньше подписантов на предыдущей странице или на предыдущей странице подпись в единственном экземпляре
                                //если на текущей странице нет подписантов, то пробуем найти слово "подпись", иначе оставляем как есть
                                TextFragment podpis_coords_average_x = get_podpis_minby_y(input_pdf_aspose_doc, i + 1);

                                try {
                                    var temp_img_coords = pdf_Document.pages.SelectMany(podpis => podpis.podpisi).Where(podpis_rect => podpis_rect.podpis_coordinates != null).Distinct().ToList();

                                    double sum = 0;

                                    for(int g = 0; g < temp_img_coords.Count; g++){
                                        sum += temp_img_coords[g].podpis_coordinates.Rectangle.Width;
                                    }

                                    average_delta_x = sum / temp_img_coords.Count;

                                    temp_img_coords = pdf_Document.pages.SelectMany(podpis => podpis.podpisi).Where(podpisant => podpisant.podpisant_coordinates != null).Distinct().ToList();

                                    sum = 0;

                                    for(int g = 0; g < temp_img_coords.Count; g++) {
                                        sum += temp_img_coords[g].podpisant_coordinates.Rectangle.Height;
                                    }

                                    average_delta_y = sum / temp_img_coords.Count;
                                } catch(System.Exception ex) {
                                }

                                double podpis_image_width = average_delta_x;
                                podpis_image_width *= 3.5;

                                double podpis_image_height = average_delta_y;
                                podpis_image_height *= 3.5;

                                if(podpis_coords_average_x != null) {
                                    try {
                                        podpis_coords_average_x.Rectangle.LLY = pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpisant_coordinates.Rectangle.LLY;

                                        Aspose.Pdf.Rectangle podpis_image_coordinates = new Aspose.Pdf.Rectangle(
                                            podpis_coords_average_x.Rectangle.LLX + (podpis_coords_average_x.Rectangle.Width / 2) - (podpis_image_width / 2),
                                            podpis_coords_average_x.Rectangle.LLY,
                                            podpis_coords_average_x.Rectangle.LLX + (podpis_coords_average_x.Rectangle.Width / 2) - (podpis_image_width / 2) + podpis_image_width,
                                            podpis_coords_average_x.Rectangle.LLY + podpis_image_height
                                        );

                                        pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_coordinates = podpis_coords_average_x;//последней подписи на предыдущей странице ставим координаты первой подписи с текущей
                                        pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_image_coordinates = podpis_image_coordinates;
                                                                                
                                        System.Drawing.Image pechat_img = System.Drawing.Image.FromFile(pechat_dlya_protocolov_file_path);

                                        Aspose.Pdf.Rectangle pechat_image_coordinates = null;

                                        pechat_image_coordinates = new Aspose.Pdf.Rectangle(
                                            pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_coordinates.Rectangle.LLX + (pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_coordinates.Rectangle.Width / 2) - ((pechat_img.Width * 0.25) / 2),
                                            pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpisant_coordinates.Rectangle.LLY - ((pechat_img.Height * 0.8) / 2),
                                            pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_coordinates.Rectangle.LLX + (pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpis_coordinates.Rectangle.Width / 2) - ((pechat_img.Width * 0.25) / 2) + (pechat_img.Width * 0.25),
                                            pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].podpisant_coordinates.Rectangle.LLY + ((pechat_img.Height * 0.8) / 2)
                                        );

                                        pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1].pechat_image_coordinates = pechat_image_coordinates;

                                        pechat_img.Dispose();
                                        pechat_img = null;
                                    } catch(System.Exception ex) {
                                    }
                                }                                
                            }
                        }
                    }

                    try {
                        Aspose.Pdf.Rectangle pechat_image_coordinates = null;

                        if((i == 0) && (podpisi.Count > 0))//если 1я страница
                        {
                            double podpisant_width = podpisi[0].podpisant_coordinates.Rectangle.Width;

                            System.Drawing.Image pechat_img = System.Drawing.Image.FromFile(pechat_organizatcii_file_path);

                            //если есть слово подпись, то относительно него получаем координаты для печати
                            if(podpisi[0].podpis_coordinates != null) {
                                pechat_image_coordinates = new Aspose.Pdf.Rectangle(
                                    podpisi[0].podpis_coordinates.Rectangle.LLX + (podpisi[0].podpis_coordinates.Rectangle.Width / 2) - ((pechat_img.Width * 0.25) / 2),
                                    podpisi[0].podpisant_coordinates.Rectangle.LLY - ((pechat_img.Height * 0.8) / 2),
                                    podpisi[0].podpis_coordinates.Rectangle.LLX + (podpisi[0].podpis_coordinates.Rectangle.Width / 2) - ((pechat_img.Width * 0.25) / 2) + (pechat_img.Width * 0.25),
                                    podpisi[0].podpisant_coordinates.Rectangle.LLY + ((pechat_img.Height * 0.8) / 2)
                                );
                            } else {
                                pechat_image_coordinates = new Aspose.Pdf.Rectangle(
                                        podpisi[0].podpis_image_coordinates.LLX + (podpisi[0].podpis_image_coordinates.Width / 2) - ((pechat_img.Width * 0.25) / 2),
                                        podpisi[0].podpis_image_coordinates.LLY + (podpisi[0].podpis_image_coordinates.Height / 2) + ((pechat_img.Height * 0.8) / 2),
                                        podpisi[0].podpis_image_coordinates.LLX + (podpisi[0].podpis_image_coordinates.Width / 2) - ((pechat_img.Width * 0.25) / 2) + (pechat_img.Width * 0.25),
                                        podpisi[0].podpis_image_coordinates.LLY + (podpisi[0].podpis_image_coordinates.Height / 2) - ((pechat_img.Height * 0.8) / 2)
                                    );
                            }

                            podpisi[0].pechat_image_coordinates = pechat_image_coordinates;

                            pechat_img.Dispose();
                            pechat_img = null;
                        } else if(podpisi.Count > 0) {
                            try {
                                S_Podpis_on_page_data podpis_on_page_data_min_y = podpisi.Where(podpis => podpis.podpis_coordinates != null).Where(podpis_rect => podpis_rect.podpis_coordinates.Rectangle != null).MinBy(podpis_min_y => podpis_min_y.podpis_coordinates.Rectangle.LLY).First();//получаем подпись по min Y из списка и даем ему координаты печати

                                double podpisant_width = podpisi[0].podpisant_coordinates.Rectangle.Width;

                                System.Drawing.Image pechat_img = System.Drawing.Image.FromFile(pechat_dlya_protocolov_file_path);

                                pechat_image_coordinates = null;

                                pechat_image_coordinates = new Aspose.Pdf.Rectangle(
                                    podpis_on_page_data_min_y.podpis_coordinates.Rectangle.LLX + (podpis_on_page_data_min_y.podpis_coordinates.Rectangle.Width / 2) - ((pechat_img.Width * 0.25) / 2),
                                    podpis_on_page_data_min_y.podpisant_coordinates.Rectangle.LLY - ((pechat_img.Height * 0.8) / 2),
                                    podpis_on_page_data_min_y.podpis_coordinates.Rectangle.LLX + (podpis_on_page_data_min_y.podpis_coordinates.Rectangle.Width / 2) - ((pechat_img.Width * 0.25) / 2) + (pechat_img.Width * 0.25),
                                    podpis_on_page_data_min_y.podpisant_coordinates.Rectangle.LLY + ((pechat_img.Height * 0.8) / 2)
                                );

                                podpis_on_page_data_min_y.pechat_image_coordinates = pechat_image_coordinates;

                                pechat_img.Dispose();
                                pechat_img = null;
                            } catch(System.Exception ex) {
                                S_Podpis_on_page_data podpis_on_page_data_min_y = podpisi.Where(podpis => podpis.podpisant_coordinates != null).Where(podpis_rect => podpis_rect.podpisant_coordinates.Rectangle != null).MinBy(podpis_min_y => podpis_min_y.podpisant_coordinates.Rectangle.LLY).First();//получаем подпись по min Y из списка и даем ему координаты печати

                                double podpisant_width = podpisi[0].podpisant_coordinates.Rectangle.Width;

                                System.Drawing.Image pechat_img = System.Drawing.Image.FromFile(pechat_dlya_protocolov_file_path);

                                pechat_image_coordinates = null;

                                pechat_image_coordinates = new Aspose.Pdf.Rectangle(
                                    podpis_on_page_data_min_y.podpisant_coordinates.Rectangle.LLX - (podpis_on_page_data_min_y.podpisant_coordinates.Rectangle.Width * 2.4),
                                    podpis_on_page_data_min_y.podpisant_coordinates.Rectangle.LLY - ((pechat_img.Height * 0.8) / 2),
                                    podpis_on_page_data_min_y.podpisant_coordinates.Rectangle.LLX - (podpis_on_page_data_min_y.podpisant_coordinates.Rectangle.Width * 2.4) + (pechat_img.Width * 0.25),
                                    podpis_on_page_data_min_y.podpisant_coordinates.Rectangle.LLY + ((pechat_img.Height * 0.8) / 2)
                                );

                                podpis_on_page_data_min_y.pechat_image_coordinates = pechat_image_coordinates;

                                pechat_img.Dispose();
                                pechat_img = null;
                            }
                        }
                    } catch(System.Exception ex) {
                    }

                    pdf_Page_Info.podpisi = podpisi;//добавляем подписи

                    pdf_Document.pages.Add(pdf_Page_Info);//добавляем инфу по странице
                }

                //ставим подписи и печати
                try{
                    for(int i = 0; i < pdf_Document.pages.Count; i++) {
                        for(int j = 0; j < pdf_Document.pages[i].podpisi.Count; j++) {
                            //ставим подписи
                            input_pdf_aspose_doc.Pages[pdf_Document.pages[i].numberOfPDFPage_].AddImage(
                                pdf_Document.pages[i].podpisi[j].podpisant.podpis_path,
                                pdf_Document.pages[i].podpisi[j].podpis_image_coordinates
                            );
                        }

                        //ставим печать
                        if((i == 0) && (pdf_Document.pages[i].podpisi.Count > 0)) {
                            if(pdf_Document.pages[i].podpisi[0].pechat_image_coordinates != null) {
                                input_pdf_aspose_doc.Pages[pdf_Document.pages[i].numberOfPDFPage_].AddImage(
                                    pechat_organizatcii_file_path,
                                    pdf_Document.pages[i].podpisi[0].pechat_image_coordinates
                                );
                            }
                        } else if(i == (pdf_Document.pages.Count - 1)) {//если последняя страница
                            if(pdf_Document.pages[i].podpisi.Count > 0) {
                                IExtremaEnumerable<S_Podpis_on_page_data> podpis_on_page_data_min_y = pdf_Document.pages[i].podpisi.Where(podpis => podpis.podpis_coordinates != null).Where(podpis_rect => podpis_rect.podpis_coordinates.Rectangle != null).MinBy(podpis_min_y => podpis_min_y.podpis_coordinates.Rectangle.LLY);//получаем подпись по min Y из списка и даем ему координаты печати

                                if(podpis_on_page_data_min_y.Count() > 0) {
                                    input_pdf_aspose_doc.Pages[pdf_Document.pages[i].numberOfPDFPage_].AddImage(
                                        pechat_dlya_protocolov_file_path,
                                        podpis_on_page_data_min_y.First().pechat_image_coordinates
                                    );
                                } else {
                                    S_Podpis_on_page_data single_podpis_on_page_data_min_y = pdf_Document.pages[i].podpisi[pdf_Document.pages[i].podpisi.Count - 1];//получаем подпись по min Y от единственной подписи

                                    input_pdf_aspose_doc.Pages[pdf_Document.pages[i].numberOfPDFPage_].AddImage(
                                        pechat_dlya_protocolov_file_path,
                                        single_podpis_on_page_data_min_y.pechat_image_coordinates
                                    );
                                }
                            } else if((pdf_Document.pages[i - 1].podpisi.Count > 0) && (pdf_Document.pages[i].podpisi.Count == 0)) {//если на текущей странице пусто, а на предыдущей есть подписанты
                                IExtremaEnumerable<S_Podpis_on_page_data> podpis_on_page_data_min_y = pdf_Document.pages[i - 1].podpisi.Where(podpis => podpis.podpis_coordinates != null).Where(podpis_rect => podpis_rect.podpis_coordinates.Rectangle != null).MinBy(podpis_min_y => podpis_min_y.podpis_coordinates.Rectangle.LLY);//получаем подпись по min Y из списка и даем ему координаты печати

                                if(podpis_on_page_data_min_y.Count() > 0) {
                                    input_pdf_aspose_doc.Pages[pdf_Document.pages[i - 1].numberOfPDFPage_].AddImage(
                                        pechat_dlya_protocolov_file_path,
                                        podpis_on_page_data_min_y.First().pechat_image_coordinates
                                    );
                                } else {
                                    S_Podpis_on_page_data single_podpis_on_page_data_min_y = pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1];//получаем подпись по min Y от единственной подписи

                                    input_pdf_aspose_doc.Pages[pdf_Document.pages[i - 1].numberOfPDFPage_].AddImage(
                                        pechat_dlya_protocolov_file_path,
                                        single_podpis_on_page_data_min_y.pechat_image_coordinates
                                    );
                                }
                            }
                        } else if(i > 1) {
                            if((pdf_Document.pages[i].podpisi.Count == 0) && (pdf_Document.pages[i - 1].podpisi.Count > 0)) {//если на текущей странице нет подписей, то на предыдущей ставим печать
                                if(pdf_Document.pages[i - 1].podpisi.Count == 1) {
                                    S_Podpis_on_page_data podpis_on_page_data_min_y = pdf_Document.pages[i - 1].podpisi[0];//получаем подпись по min Y из списка и даем ему координаты печати

                                    input_pdf_aspose_doc.Pages[pdf_Document.pages[i - 1].numberOfPDFPage_].AddImage(
                                        pechat_dlya_protocolov_file_path,
                                        podpis_on_page_data_min_y.pechat_image_coordinates
                                    );
                                } else {
                                    IExtremaEnumerable<S_Podpis_on_page_data> podpis_on_page_data_min_y = pdf_Document.pages[i - 1].podpisi.Where(podpis => podpis.podpis_coordinates != null).Where(podpis_rect => podpis_rect.podpis_coordinates.Rectangle != null).MinBy(podpis_min_y => podpis_min_y.podpis_coordinates.Rectangle.LLY);//получаем подпись по min Y из списка и даем ему координаты печати

                                    if(podpis_on_page_data_min_y.Count() > 0) {
                                        input_pdf_aspose_doc.Pages[pdf_Document.pages[i - 1].numberOfPDFPage_].AddImage(
                                            pechat_dlya_protocolov_file_path,
                                            podpis_on_page_data_min_y.First().pechat_image_coordinates
                                        );
                                    } else {
                                        S_Podpis_on_page_data single_podpis_on_page_data_min_y = pdf_Document.pages[i - 1].podpisi[pdf_Document.pages[i - 1].podpisi.Count - 1];//получаем подпись по min Y от единственной подписи

                                        input_pdf_aspose_doc.Pages[pdf_Document.pages[i - 1].numberOfPDFPage_].AddImage(
                                            pechat_dlya_protocolov_file_path,
                                            single_podpis_on_page_data_min_y.pechat_image_coordinates
                                        );
                                    }
                                }
                            }
                        }

                        if(i > 1) {//проверяем наличие начала документа по шапке, если есть, то ставим на предыдущей странице печать
                            TextFragmentAbsorber tf_check_organization = new TextFragmentAbsorber("Общество с ограниченной ответственностью \"СЕРКОНС\"");

                            input_pdf_aspose_doc.Pages[i + 1].Accept(tf_check_organization);//ищем подписанта на странице

                            if(tf_check_organization.TextFragments.Count > 0) {
                                if((pdf_Document.pages[i - 1].podpisi.Count > 0)) {
                                    S_Podpis_on_page_data podpis_on_page_data_min_y = pdf_Document.pages[i - 1].podpisi.Where(podpis => podpis.podpis_coordinates != null).Where(podpis_rect => podpis_rect.podpis_coordinates.Rectangle != null).MinBy(podpis_min_y => podpis_min_y.podpis_coordinates.Rectangle.LLY).First();//получаем подпись по min Y из списка и даем ему координаты печати

                                    input_pdf_aspose_doc.Pages[pdf_Document.pages[i - 1].numberOfPDFPage_].AddImage(
                                        pechat_dlya_protocolov_file_path,
                                        podpis_on_page_data_min_y.pechat_image_coordinates
                                    );
                                }
                            }

                            tf_check_organization = null;
                        }
                    }
                } catch (System.Exception ex){   	
                }

                //очищаем данные
                input_pdf_aspose_doc.Save(out_pdf_file);

                input_pdf_aspose_doc.Dispose();
                input_pdf_aspose_doc = null;

                for(int z = 0; z < pages.Count; z++) {
                    for(int f = 0; f < pages[z].podpisi.Count; f++) {
                        pages[z].podpisi[f].pechat_image_coordinates = null;
                        pages[z].podpisi[f].podpis_coordinates = null;
                        pages[z].podpisi[f].podpis_image_coordinates = null;
                        pages[z].podpisi[f].podpisant_coordinates = null;
                    }

                    pages[z].podpisi.Clear();
                }

                pages.Clear();
                pages = null;
            } catch(System.Exception ex) {
                string crash_message = DateTime.Now.ToString("dd.MM.yyyy HH-mm-ss") + " " + ex.Message + Environment.NewLine;

                try {
                    File.AppendAllText(outCrashFile, crash_message);
                } catch(System.Exception ex_child) {
                    try {
                        backgroundWorker1.CancelAsync();
                    } catch(System.Exception inEx) {
                    }

                    System.Environment.Exit(0);
                }

                return false;
            }

            return true;
        }

        private TextFragment get_podpis_minby_y(Aspose.Pdf.Document input_pdf_aspose_doc, int page_number) {
            TextFragmentAbsorber podpisi_coords = new TextFragmentAbsorber(@"(?i)" + @"подпись", new TextSearchOptions(true));

            //ищем подписи
            input_pdf_aspose_doc.Pages[page_number].Accept(podpisi_coords);//ищем слово "подпись" на конкретной странице

            TextFragment podpis_average_y = null;

            if(podpisi_coords.TextFragments.Count > 0) {
                List<TextFragment> temp_text_fragments = new List<TextFragment>();

                for(int i = 0; i < podpisi_coords.TextFragments.Count; i++) {
                    temp_text_fragments.Add(podpisi_coords.TextFragments[i+1]);
                }

                podpis_average_y = temp_text_fragments.MinBy(text_fragment_y => text_fragment_y.Rectangle.LLY).First();

                return podpis_average_y;
            } else {
                return podpis_average_y;
            }
        }

        private int get_podpisi_count(Aspose.Pdf.Document input_pdf_aspose_doc, int page_number) {
            TextFragmentAbsorber podpisi_coords = new TextFragmentAbsorber(@"(?i)" + @"подпись", new TextSearchOptions(true));

            //ищем подписи
            input_pdf_aspose_doc.Pages[page_number].Accept(podpisi_coords);//ищем слово "подпись" на конкретной странице

            return podpisi_coords.TextFragments.Count;
        }

        private List<S_Podpis_on_page_data> get_podpis_on_page_data(Aspose.Pdf.Document input_pdf_aspose_doc, string input_pdf_file_path, S_Podpisant podpisant, int page_number) {
            List<S_Podpis_on_page_data> podpisi = new List<S_Podpis_on_page_data>();
            List<S_Podpis_on_page_data> temp_podpisi = null;

            string tf_podpisanti_mask = "";

            int number_of_podpisant_mask_combination = -1;

            if(podpisant.familia != null && podpisant.imia != null && podpisant.otchestvo != null) {
                tf_podpisanti_mask = podpisant.familia + @"[ ]*" + podpisant.imia + @"[ ]*" + podpisant.otchestvo;

                number_of_podpisant_mask_combination = 1;
            } else if(podpisant.familia != null && podpisant.imia != null && podpisant.otchestvo == null) {
                tf_podpisanti_mask = podpisant.familia + @"[ ]*" + podpisant.imia;

                number_of_podpisant_mask_combination = 2;
            } else if(podpisant.familia != null && podpisant.imia == null && podpisant.otchestvo == null) {
                tf_podpisanti_mask = podpisant.familia + @"[ ]{1}";

                number_of_podpisant_mask_combination = 3;
            }

            temp_podpisi = get_podpisi_from_text_fragment(input_pdf_aspose_doc, input_pdf_file_path, podpisant, page_number, tf_podpisanti_mask, true);

            if(temp_podpisi.Count > 0) {//проверка всех вариантов масок (1, 2, 3)
                for(int k = 0; k < temp_podpisi.Count; k++) {
                    podpisi.Add(temp_podpisi[k]);
                }
            } else if((temp_podpisi.Count == 0) && (number_of_podpisant_mask_combination == 3)) {//если не нашли, то пробуем без маски (3)
                temp_podpisi = get_podpisi_from_text_fragment(input_pdf_aspose_doc, input_pdf_file_path, podpisant, page_number, tf_podpisanti_mask, false);

                if(temp_podpisi.Count > 0) {
                    for(int k = 0; k < temp_podpisi.Count; k++) {
                        podpisi.Add(temp_podpisi[k]);
                    }
                }
            }

            //проверяем сокращения (1)
            if((temp_podpisi.Count == 0) && (number_of_podpisant_mask_combination == 1)) {
                //пробуем фамилия + имя
                //имя справа (1)
                tf_podpisanti_mask = podpisant.familia + @"[ ]*" + podpisant.imia;

                temp_podpisi = get_podpisi_from_text_fragment(input_pdf_aspose_doc, input_pdf_file_path, podpisant, page_number, tf_podpisanti_mask, true);

                if(temp_podpisi.Count > 0) {
                    for(int k = 0; k < temp_podpisi.Count; k++) {
                        podpisi.Add(temp_podpisi[k]);
                    }
                } else {//если не нашли, пробуем поставить имя слева(1)
                    tf_podpisanti_mask = podpisant.imia[0] + @"[ .]*" + podpisant.familia;

                    temp_podpisi = get_podpisi_from_text_fragment(input_pdf_aspose_doc, input_pdf_file_path, podpisant, page_number, tf_podpisanti_mask, true);

                    if(temp_podpisi.Count > 0) {
                        for(int k = 0; k < temp_podpisi.Count; k++) {
                            podpisi.Add(temp_podpisi[k]);
                        }
                    }
                }

                //берем сокращение имени и отчества, ставим их слева и справа
                tf_podpisanti_mask = podpisant.familia + @"[ ]*" + podpisant.imia[0] + @"[ .]*" + podpisant.otchestvo[0] + @"[ .]*";

                temp_podpisi = get_podpisi_from_text_fragment(input_pdf_aspose_doc, input_pdf_file_path, podpisant, page_number, tf_podpisanti_mask, true);

                if(temp_podpisi.Count > 0) {
                    for(int k = 0; k < temp_podpisi.Count; k++) {
                        podpisi.Add(temp_podpisi[k]);
                    }
                } else {//если не нашли, пробуем поставить слева(1)
                    tf_podpisanti_mask = podpisant.imia[0] + @"[ .]*" + podpisant.otchestvo[0] + @"[ .]*" + podpisant.familia;

                    temp_podpisi = get_podpisi_from_text_fragment(input_pdf_aspose_doc, input_pdf_file_path, podpisant, page_number, tf_podpisanti_mask, true);

                    if(temp_podpisi.Count > 0) {
                        for(int k = 0; k < temp_podpisi.Count; k++) {
                            podpisi.Add(temp_podpisi[k]);
                        }
                    }
                }
            }

            //проверяем сокращения (2)
            if((temp_podpisi.Count == 0) && (number_of_podpisant_mask_combination == 2)) {
                //берем сокращение имени и отчества, ставим их слева и справа
                tf_podpisanti_mask = podpisant.familia + @"[ ]*" + podpisant.imia[0] + @"[ .]*";

                temp_podpisi = get_podpisi_from_text_fragment(input_pdf_aspose_doc, input_pdf_file_path, podpisant, page_number, tf_podpisanti_mask, true);

                if(temp_podpisi.Count > 0) {
                    for(int k = 0; k < temp_podpisi.Count; k++) {
                        podpisi.Add(temp_podpisi[k]);
                    }
                } else {//если не нашли, пробуем поставить слева (2)
                    tf_podpisanti_mask = podpisant.imia[0] + @"[ ]*[.]{1}" + @"\w*\W*" + podpisant.familia;

                    temp_podpisi = get_podpisi_from_text_fragment(input_pdf_aspose_doc, input_pdf_file_path, podpisant, page_number, tf_podpisanti_mask, true);

                    if(temp_podpisi.Count > 0) {
                        for(int k = 0; k < temp_podpisi.Count; k++) {
                            podpisi.Add(temp_podpisi[k]);
                        }
                    }
                }
            }

            return podpisi;
        }

        public List<S_Podpis_on_page_data> get_podpisi_from_text_fragment(Aspose.Pdf.Document input_pdf_aspose_doc, string input_pdf_file_path, S_Podpisant podpisant, int page_number, string mask, bool is_regular_expression) {
            List<S_Podpis_on_page_data> podpisi = new List<S_Podpis_on_page_data>();

            TextSearchOptions ts_options = null;

            if(is_regular_expression == true) {
                ts_options = new TextSearchOptions(true);//включаем регулярные выражения
            } else {
                ts_options = new TextSearchOptions(false);//включаем регулярные выражения
            }

            TextFragmentAbsorber tf_podpisanti = null;

            tf_podpisanti = new TextFragmentAbsorber(mask);

            tf_podpisanti.TextSearchOptions = ts_options;

            input_pdf_aspose_doc.Pages[page_number].Accept(tf_podpisanti);//ищем подписанта на странице

            TextFragmentAbsorber podpisi_coords = new TextFragmentAbsorber(@"(?i)" + @"подпись", new TextSearchOptions(true));

            //ищем подписи
            input_pdf_aspose_doc.Pages[page_number].Accept(podpisi_coords);//ищем слово "подпись" на всех страницах

            if(tf_podpisanti.TextFragments.Count > 0) {
                for(int k = 0; k < tf_podpisanti.TextFragments.Count; k++) {
                    S_Podpis_on_page_data podpis_on_page_data = new S_Podpis_on_page_data();

                    TextFragmentAbsorber absorber = new TextFragmentAbsorber(podpisant.familia);//опция для установки поиска внутри recta на предмет фамилии

                    absorber.TextSearchOptions.Rectangle = tf_podpisanti.TextFragments[k + 1].Rectangle;

                    input_pdf_aspose_doc.Pages[page_number].Accept(absorber);//ищем подписанта на странице

                    TextFragment podpis = get_podpis_coordinates_by_podpisant_coordinates_aspose(absorber.TextFragments[1], podpisi_coords);

                    podpis_on_page_data.input_pdf_file_path = input_pdf_file_path;
                    podpis_on_page_data.numberOfPDFPage = page_number;
                    podpis_on_page_data.podpisant = podpisant;
                    podpis_on_page_data.podpisant_coordinates = tf_podpisanti.TextFragments[k + 1];
                    podpis_on_page_data.podpis_coordinates = podpis;

                    podpisi.Add(podpis_on_page_data);

                    //чистим память
                    absorber = null;
                }
            }

            podpisi_coords = null;

            return podpisi;
        }

        public TextFragment get_podpis_coordinates_by_podpisant_coordinates_aspose(TextFragment podpisant_coordinates, TextFragmentAbsorber podpisi_coordinates) {
            TextFragment podpis_coordinates = null;

            //сопоставляем подписанта и подпись
            if(podpisi_coordinates.TextFragments.Count > 0) {
                double delta_y_min_to_podpis_from_podpisant = 0;

                int index_min_y_podpis = -1;

                for(int i = 0; i < podpisi_coordinates.TextFragments.Count; i++) {
                    double temp_y_delta = Math.Abs(podpisant_coordinates.Rectangle.URY - podpisi_coordinates.TextFragments[i + 1].Rectangle.URY);

                    if(i == 0) {
                        delta_y_min_to_podpis_from_podpisant = temp_y_delta;

                        index_min_y_podpis = i;
                    } else if(temp_y_delta < delta_y_min_to_podpis_from_podpisant) {
                        delta_y_min_to_podpis_from_podpisant = temp_y_delta;

                        index_min_y_podpis = i;
                    }
                }

                podpis_coordinates = podpisi_coordinates.TextFragments[index_min_y_podpis + 1];
            }

            return podpis_coordinates;
        }

        public struct S_Pdf_Document {
            public List<S_Pdf_Page_Info> pages;//страницы в документе
            public string input_pdf_file_path;//путь к родительскому документу
        }

        public struct S_Pdf_Page_Info {
            public List<S_Podpis_on_page_data> podpisi;//подписи
            public int numberOfPDFPage_;//номер страницы в документе
            public string input_pdf_file_path;//путь к родительскому документу
        }

        public struct S_Podpisant {
            public List<Regex> maska;//маски для поиска ФИО
            //public System.Drawing.Image podpis_img;//картина подписи
            public string fio;//ФИО
            public string familia;//фамилия
            public string imia;//имя
            public string otchestvo;//отчество
            public string podpis_path;//путь к файлу с подписью(JPG)
        }

        public class S_Podpis_on_page_data {
            public S_Podpisant podpisant;//данные по подписанту
            public TextFragment podpisant_coordinates;//координаты подписанта(место где его ФИО написано)
            public TextFragment podpis_coordinates;//координаты подписи(место куда ставится подпись подписанта)
            public Aspose.Pdf.Rectangle podpis_image_coordinates;//координаты подставляемого изображения подписи
            public Aspose.Pdf.Rectangle pechat_image_coordinates;//координаты подставляемого изображения печати
            public int numberOfPDFPage;//номер страницы в документе
            public string input_pdf_file_path;//путь к родительскому документу
        }

        public List<S_Podpisant> get_podpisanti_data(string podpisanti_papka) {
            List<S_Podpisant> podpisanti = new List<S_Podpisant>();//подписанты

            List<string> podpisanti_files = getImageFiles(podpisanti_papka);//список подписантов(отмасштабированные подписи)

            //собираем подписантов
            for (int i = 0; i < podpisanti_files.Count; i++) {
                S_Podpisant podpisant = new S_Podpisant();

                podpisant.podpis_path = podpisanti_files[i];//путь к подписи

                podpisant.fio = Path.GetFileNameWithoutExtension(podpisanti_files[i]);//ФИО

                zapolnit_fio_podpisanta_razdelno(ref podpisant, Path.GetFileNameWithoutExtension(podpisanti_files[i]));//разбиваем фио на части             

                List<string> fio_maska_string_list = razdelit_stroky(Path.GetFileNameWithoutExtension(podpisanti_files[i]), 2);//делим строку с ФИО на части для маски

                podpisant.maska = stroki_v_regexi(fio_maska_string_list);//пишем маску

                podpisanti.Add(podpisant);//добавляем подписанта
            }

            return podpisanti;
        }

        private void delete_folder(string input_folder_path) {
            do {
                try {
                    if(System.IO.Directory.Exists(input_folder_path)) {
                        System.IO.Directory.Delete(input_folder_path, true);
                    }
                } catch(System.Exception ex) {
                    string crash_message = DateTime.Now.ToString("dd.MM.yyyy HH-mm-ss") + " " + ex.Message + Environment.NewLine;

                    try {
                        File.AppendAllText(outCrashFile, crash_message);
                    } catch(System.Exception ex_child) {
                        try {
                            backgroundWorker1.CancelAsync();
                        } catch(System.Exception inEx) {
                        }

                        System.Environment.Exit(0);
                    }
                }
            } while(System.IO.Directory.Exists(input_folder_path) == true);
        }

        private void result_print(string file_path, bool is_podpis_i_pechati_prostavleni) {
            try {
                string done_files = outFilePath + "\\Done_files.txt";
                string error_files = outFilePath + "\\Error_files.txt";

                if(is_podpis_i_pechati_prostavleni == true) {
                    File.AppendAllText(done_files, file_path + Environment.NewLine);
                } else {
                    File.AppendAllText(error_files, file_path + Environment.NewLine);
                }
            } catch(System.Exception ex) {
            }
        }

        private void result_print_old(S_Pdf_Document pdf_document, bool is_podpis_i_pechati_prostavleni, System.IO.StreamWriter done_files, System.IO.StreamWriter error_files) {
            if(is_podpis_i_pechati_prostavleni == true) {
                done_files.WriteLine(pdf_document.input_pdf_file_path);
            } else {
                error_files.WriteLine(pdf_document.input_pdf_file_path);
            }
        }

        private List<string> getImageFiles(string inputFolder) {
            List<string> fileNames = new List<string>();

            try {
                //берем список имен файлов в папке
                string[] fNames = System.IO.Directory.GetFiles(inputFolder, "*", SearchOption.AllDirectories);

                //убираем архивные файлы
                for(int i = 0; i < fNames.Length; i++) {
                    string extension = Path.GetExtension(fNames[i]);

                    if(extension == ".jpg" || extension == ".jpeg" || extension == ".JPG" || extension == ".JPEG" ||
                        extension == ".png" || extension == ".PNG" || extension == ".bmp" || extension == ".BMP" ||
                        extension == ".ico" || extension == ".ICO"
                        ) {
                        fileNames.Add(fNames[i]);
                    }
                }
            } catch(System.Exception ex) {
                return fileNames;
            }

            return fileNames;
        }

        public void create_folder(string folder_path) {
            do {
                System.IO.Directory.CreateDirectory(folder_path);
            } while(System.IO.Directory.Exists(folder_path) == false);
        }

        public string guid_name() {
            return Guid.NewGuid().ToString();
        }

        private void clear_pdf(string input_file_path) {
            //список для чистки
            List<string> replace_text_list = new List<string>() {
                "Created with an evaluation copy of Aspose.Words. To discover the full versions of our APIs please visit: https://products.aspose.com/words/",
                "Evaluation Only. Created with Aspose.Words. Copyright 2003-2019 Aspose Pty Ltd.",
                "Evaluation Warning: The document was created with Spire.Doc for .NET.",
                "Evaluation Warning : The document was created with Spire.PDF for .NET.",
                "This document was truncated here because it was created in the Evaluation Mode."
            };

            //делаем копию, чтобы считать обновленный контент
            string imia_file_update = Path.GetFileNameWithoutExtension(input_file_path);//имя входного файла
            string out_pdf_file_update = outFilePath + "\\" + imia_file_update + " " + DateTime.Now.ToString("dd.MM.yyyy HH-mm-ss") + ".pdf";

            for(int i = 0; i < replace_text_list.Count; i++) {
                VerySimpleReplaceText(input_file_path, out_pdf_file_update, replace_text_list[i], "");
            }
        }

        public void VerySimpleReplaceText(string OrigFile, string ResultFile, string origText, string replaceText) {
            using(iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(OrigFile)) {
                for(int i = 1; i <= reader.NumberOfPages; i++) {
                    byte[] contentBytes = reader.GetPageContent(i);
                    string contentString = PdfEncodings.ConvertToString(contentBytes, iTextSharp.text.pdf.PdfObject.TEXT_PDFDOCENCODING);
                    contentString = contentString.Replace(origText, replaceText);
                    reader.SetPageContent(i, PdfEncodings.ConvertToBytes(contentString, iTextSharp.text.pdf.PdfObject.TEXT_PDFDOCENCODING));
                }
                new PdfStamper(reader, new FileStream(ResultFile, FileMode.Create, FileAccess.Write)).Close();
            }
        }

        private void zapolnit_fio_podpisanta_razdelno(ref S_Podpisant podpisant, string fio) {
            List<string> fio_split = Path.GetFileNameWithoutExtension(fio).Split(' ').ToList();

            if(fio_split.Count == 3) {
                podpisant.familia = fio_split[0];
                podpisant.imia = fio_split[1];
                podpisant.otchestvo = fio_split[2];
            } else if(fio_split.Count == 2) {
                podpisant.familia = fio_split[0];
                podpisant.imia = fio_split[1];
            } else if(fio_split.Count == 1) {
                podpisant.familia = fio_split[0];
            }
        }

        public float delta_x_calculate(TRectAndText first_point, TRectAndText second_point) {
            float delta = Math.Abs((long)second_point.Rect.Left - first_point.Rect.Left);

            return delta;
        }

        public float delta_y_calculate(TRectAndText first_point, TRectAndText second_point) {
            float delta = Math.Abs((long)second_point.Rect.Top - first_point.Rect.Bottom);

            return delta;
        }

        public double distance_calculate(TRectAndText first_point, TRectAndText second_point) {
            double distance = Math.Sqrt(
                Math.Pow(Math.Abs((long)first_point.Rect.Left - second_point.Rect.Left), 2) +
                Math.Pow(Math.Abs((long)first_point.Rect.Top - second_point.Rect.Top), 2)
            );

            return distance;
        }

        public enum ocrMethod { TESSERACT = 1, ABBY = 2, NONE = 3 };

        public Tuple<bool, int> is_naiden_podpisant_v_texte(S_Podpisant podpisant, string text) {
            Tuple<bool, int> is_naiden_podpisant = new Tuple<bool, int>(false, 0);

            Regex podpisant_regex = new Regex(@"\w*" + podpisant.familia + @"\w*", RegexOptions.IgnoreCase);

            MatchCollection podpisant_matches = podpisant_regex.Matches(text);

            if(podpisant_matches.Count > 0) {
                is_naiden_podpisant = new Tuple<bool, int>(true, podpisant_matches.Count);
            } else {
                is_naiden_podpisant = new Tuple<bool, int>(false, 0);
            }

            return is_naiden_podpisant;
        }

        public bool is_naideno_v_texte(string input_string, string text) {
            Regex regex = new Regex(@"\w*" + input_string + @"\w*", RegexOptions.IgnoreCase);

            MatchCollection matches = regex.Matches(text);

            if(matches.Count > 0) {
                return true;
            } else {
                return false;
            }
        }

        public void saveErrorFile(Exception error, string fileName) {
            //пишем список оставшихся файлов для распознания, если упало
            string currentPath = System.IO.Directory.GetCurrentDirectory();

            string outCrashFolder = currentPath + "\\[Crash]Log";

            System.IO.Directory.CreateDirectory(outCrashFolder);

            if((files_spisok.Count > 0) && (curIndFileNames > 0)) {
                files_spisok.RemoveRange(0, curIndFileNames - 1);//откатываемся на один назад, чтобы затереть проблемный
            }

            System.IO.File.WriteAllLines(outCrashFolder + "\\[Crash]Remaining_files.txt", files_spisok.ToArray());
            System.IO.File.AppendAllText(outCrashFolder + "\\[Crash]Log.txt", fileName + Environment.NewLine);//выводим ошибку в файл, если файл создан, то дописываем в конец
            System.IO.File.AppendAllText(outCrashFolder + "\\[Crash]Errors.txt", error.Message + Environment.NewLine);//выводим ошибку в файл, если файл создан, то дописываем в конец
        }

        public List<string> getPdfFiles(string inputFolder) {
            List<string> fileNames = new List<string>();

            {
                //берем список имен файлов в папке
                string[] fNames = System.IO.Directory.GetFiles(inputFolder, "*", SearchOption.AllDirectories);

                //убираем архивные файлы
                for(int i = 0; i < fNames.Length; i++) {
                    string extension = Path.GetExtension(fNames[i]);

                    if(extension == ".pdf") {
                        fileNames.Add(fNames[i]);
                    }
                }
            }

            return fileNames;
        }

        public List<string> geJpgFiles(string inputFolder) {
            List<string> fileNames = new List<string>();

            {
                //берем список имен файлов в папке
                string[] fNames = System.IO.Directory.GetFiles(inputFolder, "*", SearchOption.AllDirectories);

                //убираем архивные файлы
                for(int i = 0; i < fNames.Length; i++) {
                    string extension = Path.GetExtension(fNames[i]);

                    if(extension == ".jpg") {
                        fileNames.Add(fNames[i]);
                    }
                }
            }

            return fileNames;
        }

        List<string> razdelit_stroky(string stroka, int delimeter) {
            List<string> razdelennaya_stroka = new List<string>();

            if(stroka == String.Empty) {//если пустая строка - возвращаем пустой контейнер
                razdelennaya_stroka.Clear();

                return razdelennaya_stroka;
            }

            int dlina_stroki = stroka.Length;//длина строки

            if(dlina_stroki >= delimeter) {//если можно поделить с заданным шагом
                int dlina_podstroki = 0;//длина подстроки после деления
                int ostavshayasia_dlina_stroki_posle_delenia = 0;//оставшаяся часть строки(неподеленная)

                try {
                    dlina_podstroki = dlina_stroki / delimeter;
                } catch(System.Exception ex) {
                }

                try {
                    ostavshayasia_dlina_stroki_posle_delenia = dlina_stroki - ((dlina_stroki / delimeter) * delimeter);//оставшаяся длины фио
                } catch(System.Exception ex) {
                }

                if(dlina_podstroki == 0) {//если не поделилось, кладём всю строку
                    razdelennaya_stroka.Add(stroka);

                    return razdelennaya_stroka;
                } else {
                    for(int i = 0; i < delimeter; i++) {
                        string podstroka = stroka.Substring(i * dlina_podstroki, dlina_podstroki);

                        razdelennaya_stroka.Add(podstroka);
                    }
                }

                //добавляем оставшуюся часть строки
                if(ostavshayasia_dlina_stroki_posle_delenia != 0) {
                    string ostavshayasia_chast_stroki_posle_delenia = stroka.Substring(dlina_podstroki * delimeter - 1, ostavshayasia_dlina_stroki_posle_delenia);

                    razdelennaya_stroka.Add(ostavshayasia_chast_stroki_posle_delenia);
                }

                return razdelennaya_stroka;
            } else {
                razdelennaya_stroka.Add(stroka);

                return razdelennaya_stroka;
            }
        }

        List<Regex> stroki_v_regexi(List<string> stroki) {
            List<Regex> regexi = new List<Regex>();

            for(int i = 0; i < stroki.Count; i++) {
                Regex regex = new Regex(@"\w*" + stroki[i] + @"\w*", RegexOptions.Singleline);
                regexi.Add(regex);
            }

            return regexi;
        }

        private void docButton_Click(object sender, EventArgs e) {
            if(openFileDialog.ShowDialog() == DialogResult.OK) {
                input_document_file = openFileDialog.FileName;//берем файлик со списком доков для распознания
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            input_document_file = "";
        }
    }
}