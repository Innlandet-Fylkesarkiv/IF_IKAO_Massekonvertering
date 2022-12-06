using ImageMagick;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace IF_IKAO_Massekonvertering {
    public partial class Form1 : Form {
        delegate void SetTextCallback(string text, Label label);

        public List<string> JPGList;
        public List<string> TIFList;
        public List<string> GIFList;
        public List<string> PNGList;
        public List<string> FolderList;
        public List<string> MiscList;

        public string InFolder = "INN";
        public string OutFolder = "OUT";
        public string errorsCaught = "\nFiles with errors:";

        public Form1() {
            InitializeComponent();
            JPGList = new List<string>();
            TIFList = new List<string>();
            GIFList = new List<string>();
            PNGList = new List<string>();
            FolderList = new List<string>();
            MiscList = new List<string>();

            SetText(Path.GetFullPath(InFolder) + "\n\nElementer i Inn mappen:", PathAndInfoBox);

            //Sets up and checks folders
            if (!Directory.Exists(InFolder))
                Directory.CreateDirectory(InFolder);
            else if(Directory.GetDirectories(InFolder).Length != 0 || Directory.GetFiles(InFolder).Length != 0) {
                if (Directory.GetDirectories(InFolder).Length != 0)
                    foreach (string folder in Directory.GetDirectories(InFolder).ToList())
                        FolderInLoop(folder);
                if (Directory.GetFiles(InFolder).Length != 0)
                    foreach (string file in Directory.GetFiles(InFolder).ToList())
                        FileInCheck(file);
            }
            if (!Directory.Exists(OutFolder))
                Directory.CreateDirectory(OutFolder);

            UpdateInScreen();

            //Sets up the watcher
            var watcher = new FileSystemWatcher(InFolder);
            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;

            watcher.Changed += OnChanged;
            watcher.Created += OnCreated;
            watcher.Deleted += OnDeleted;
            watcher.Renamed += OnRenamed;
            watcher.Error += OnError;
            watcher.IncludeSubdirectories = true;
            watcher.EnableRaisingEvents = true;
        }

        private void StartConversion_Click(object sender, EventArgs e) {
            if (!Directory.Exists(InFolder) || !Directory.Exists(OutFolder)) {
                SetText("Enten INN mappe eller UT mappe mangler.\nVennligst restart programmet for å gjenskape disse."
                    , ConverterInfoBox);
                return;
            }
            if (Directory.GetDirectories(InFolder).Length == 0 && Directory.GetFiles(InFolder).Length == 0) {
                SetText("Vennligst mat INN mappen med noe før du forsøker å konvertere.", ConverterInfoBox);
                return;
            }
            if (Directory.GetDirectories(OutFolder).Length != 0 || Directory.GetFiles(OutFolder).Length != 0) {
                SetText("Vennligst tøm UT mappen før du forsøker å konvertere.", ConverterInfoBox);
                return;
            }
            if (FileTypeConvertTo.SelectedItem == null) {
                SetText("Vennligst velg et format å konvertere til.", ConverterInfoBox);
                return;
            }
            StartConversion.Enabled = false;
            ConverterInfoBox.Visible = false;
            ProgressBarWidget.Visible = true;
            CurrentTaskBox.Visible = true;
            this.Refresh();

            int totalFilesAndFolders = JPGList.Count + TIFList.Count + PNGList.Count + 
                                        GIFList.Count + FolderList.Count + MiscList.Count;
            int numFilesAndFoldersCompleted = 0;

            //Folders
            foreach(string fullPath in FolderList.ToList()) {
                string folder = fullPath.Substring(fullPath.IndexOf(@"\") + 1); //Gets everything after \
                string path = OutFolder + @"\";
                while (folder.Contains(@"\")) {
                    path += folder.Substring(0, folder.IndexOf(@"\"));
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    path += @"\";
                    folder = folder.Substring(folder.IndexOf(@"\") + 1);
                }
                path += folder;
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                numFilesAndFoldersCompleted++;
                int progress = (int)Math.Round((double)(1000 * numFilesAndFoldersCompleted) / (double)totalFilesAndFolders);
                ProgressBarWidget.Value = progress;
                SetText("Kopierer mappe:\n\t" + path, CurrentTaskBox);
                this.Refresh();
            }

            //Miscellaneous 
            numFilesAndFoldersCompleted = CopyFiles(MiscList, totalFilesAndFolders, numFilesAndFoldersCompleted);

            //Convertions and copies
            string convertTo = FileTypeConvertTo.GetItemText(FileTypeConvertTo.SelectedItem);
            switch (convertTo) {
                case "JPG": numFilesAndFoldersCompleted = CopyFiles(JPGList, totalFilesAndFolders, numFilesAndFoldersCompleted); break;
                case "TIF": numFilesAndFoldersCompleted = CopyFiles(TIFList, totalFilesAndFolders, numFilesAndFoldersCompleted); break;
                case "PNG": numFilesAndFoldersCompleted = CopyFiles(PNGList, totalFilesAndFolders, numFilesAndFoldersCompleted); break;
                case "GIF": numFilesAndFoldersCompleted = CopyFiles(GIFList, totalFilesAndFolders, numFilesAndFoldersCompleted); break;
                default: break;
            }

            if (convertTo != "JPG")
                numFilesAndFoldersCompleted = ConvertFiles(JPGList, convertTo, totalFilesAndFolders, numFilesAndFoldersCompleted);
            if (convertTo != "TIF")
                numFilesAndFoldersCompleted = ConvertFiles(TIFList, convertTo, totalFilesAndFolders, numFilesAndFoldersCompleted);
            if (convertTo != "PNG")
                numFilesAndFoldersCompleted = ConvertFiles(PNGList, convertTo, totalFilesAndFolders, numFilesAndFoldersCompleted);
            if (convertTo != "GIF")
                numFilesAndFoldersCompleted = ConvertFiles(GIFList, convertTo, totalFilesAndFolders, numFilesAndFoldersCompleted);

            SetText("", CurrentTaskBox);
            CurrentTaskBox.Visible = false;
            ProgressBarWidget.Visible = false;
            ConverterInfoBox.Visible = true;
            SetText("Fullført, overførte " + FolderList.Count + " mapper og " + 
                        (numFilesAndFoldersCompleted - FolderList.Count) + " filer.\n\nError:" + errorsCaught, ConverterInfoBox);
            StartConversion.Enabled = true;
        }

        /**Copies all files and sends them to the correct folders
         * List<string> copy = List of files to copy
         * returns an updated number of files completed
         */ 
        private int CopyFiles (List<string> copy, int totalFilesAndFolders, int numFilesAndFoldersCompleted) {
            int numRecursions = 0;
            string lastFileCopied;
            foreach (string fullPath in copy.ToList()) {
                string file = fullPath.Substring(fullPath.IndexOf(@"\") + 1);
                string path = OutFolder + @"\";
                File.Copy(fullPath, path + file);
                numFilesAndFoldersCompleted++;
                int progress = (int)Math.Round((double)(1000 * numFilesAndFoldersCompleted) / (double)totalFilesAndFolders);
                ProgressBarWidget.Value = progress;
                SetText("Kopierer fil:\n\t" + path + "\n\t\t" + file, CurrentTaskBox);
                this.Refresh();
                numRecursions++;
                lastFileCopied = fullPath;
            }
            return numFilesAndFoldersCompleted;
        }

        /**Converts all files and sends them to the correct folders
         * List<string> convertFrom = List of files to convert
         * string convertTo = file extension to convert to, such as PDF, JPG, etc.
         * returns an updated number of files completed
         */
        private int ConvertFiles (List<string> convertFrom, string convertTo, int totalFilesAndFolders, int numFilesAndFoldersCompleted) {
            bool merge = false;
            switch (convertTo) {
                case "PDF": convertTo = ".pdf"; break;
                case "JPG": convertTo = ".jpg"; break;
                case "TIF": convertTo = ".tif"; break;
                case "GIF": convertTo = ".gif"; break;
                case "PNG": convertTo = ".png"; break;
                case "Merge to PDF": convertTo = ".pdf"; merge = true; break;
            }

            if (merge) {
                List<string> filesInFolder = new List<string>();
                string currentDir = "";
                int numMergeFiles = 1;
                foreach (string fullPath in convertFrom.ToList()) {
                    string file = fullPath.Substring(fullPath.IndexOf(@"\") + 1);
                    string path = OutFolder + @"\";
                    while (file.Contains(@"\")) {
                        path += file.Substring(0, file.IndexOf(@"\"));
                        if (!Directory.Exists(path))
                            Directory.CreateDirectory(path);
                        path += @"\";
                        file = file.Substring(file.IndexOf(@"\") + 1);
                    }
                    if (currentDir == path && filesInFolder.Count <= 25) {
                        filesInFolder.Add(fullPath);
                    }
                    else {
                        if (filesInFolder.Count > 0) {
                            numFilesAndFoldersCompleted = mergeToPDF(filesInFolder, convertTo, totalFilesAndFolders, numFilesAndFoldersCompleted, numMergeFiles.ToString());
                            numMergeFiles++;
                        }
                        filesInFolder.Clear();
                        filesInFolder.Add(fullPath);
                        currentDir = path;
                    }
                }
                if (filesInFolder.Count > 0)
                    numFilesAndFoldersCompleted = mergeToPDF(filesInFolder, convertTo, totalFilesAndFolders, numFilesAndFoldersCompleted, numMergeFiles.ToString());
                return numFilesAndFoldersCompleted;
            }
            
            foreach(string fullPath in convertFrom.ToList()) {
                string file = fullPath.Substring(fullPath.IndexOf(@"\") + 1);
                string path = OutFolder + @"\";
                while (file.Contains(@"\")) {
                    path += file.Substring(0, file.IndexOf(@"\"));
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    path += @"\";
                    file = file.Substring(file.IndexOf(@"\") + 1);
                }
                try {
                    using (var image = new MagickImage(fullPath)) {
                        image.Write(path + Path.GetFileNameWithoutExtension(file) + convertTo);
                    }
                }
                catch {
                    errorsCaught += "\n\t" + path + "\n\t\t" + file;
                    if (File.Exists(fullPath))
                        File.Copy(fullPath, path + file, false);
                }
                numFilesAndFoldersCompleted++;
                int progress = (int)Math.Round((double)(1000 * numFilesAndFoldersCompleted) / (double)totalFilesAndFolders);
                ProgressBarWidget.Value = progress;
                if (errorsCaught == "\nFiles with errors:")
                    SetText("Konverterer fil:\n\t)" + path + "\n\t\t" + file, CurrentTaskBox);
                else
                    SetText("Konverterer fil:\n\t" + path + "\n\t\t" + file + "\n\n" + errorsCaught, CurrentTaskBox);
                this.Refresh();
            }
            return numFilesAndFoldersCompleted;
        }

        private int mergeToPDF(List<string> convertFrom, string convertTo, int totalFilesAndFolders, int numFilesAndFoldersCompleted, string nameAppend) {
            string errorsCaught = "\nFiles with errors:";
            using (var images = new MagickImageCollection()) {
                string fullPath = "";
                foreach (string realPath in convertFrom.ToList()) {
                    string filen = fullPath.Substring(fullPath.IndexOf(@"\") + 1);
                    string pathn = OutFolder + @"\";
                    while (filen.Contains(@"\")) {
                        pathn += filen.Substring(0, filen.IndexOf(@"\"));
                        if (!Directory.Exists(pathn))
                            Directory.CreateDirectory(pathn);
                        pathn += @"\";
                        filen = filen.Substring(filen.IndexOf(@"\") + 1);
                    }
                    images.Add(new MagickImage(realPath));
                    fullPath = realPath;
                    numFilesAndFoldersCompleted++;
                    int progress = (int)Math.Round((double)(1000 * numFilesAndFoldersCompleted) / (double)totalFilesAndFolders);
                    ProgressBarWidget.Value = progress;
                    if (errorsCaught == "\nFiles with errors:")
                        SetText("Legger til fil:\n\t)" + pathn + "\n\t\t" + filen, CurrentTaskBox);
                    else
                        SetText("Legger til fil:\n\t" + pathn + "\n\t\t" + filen + errorsCaught, CurrentTaskBox);
                    this.Refresh();
                }
                string file = fullPath.Substring(fullPath.IndexOf(@"\") + 1);
                string path = OutFolder + @"\";
                while (file.Contains(@"\")) {
                    path += file.Substring(0, file.IndexOf(@"\"));
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    path += @"\";
                    file = file.Substring(file.IndexOf(@"\") + 1);
                }
                if (errorsCaught == "\nFiles with errors:")
                    SetText("Samler fil:\n\t)" + path + "\n\t\t" + file, CurrentTaskBox);
                else
                    SetText("Samler fil:\n\t" + path + "\n\t\t" + file + errorsCaught, CurrentTaskBox);
                this.Refresh();
                try {
                    images.Write(path + "Merged file " + nameAppend + convertTo);
                }
                catch {
                    errorsCaught += "\n\t" + path + "\n\t\t" + file;
                    if (File.Exists(fullPath))
                        File.Copy(fullPath, path + file, false);
                }
                if (errorsCaught == "\nFiles with errors:")
                    SetText("Konverterer fil:\n\t)" + path + "\n\t\t" + file, CurrentTaskBox);
                else
                    SetText("Konverterer fil:\n\t" + path + "\n\t\t" + file + errorsCaught, CurrentTaskBox);
                this.Refresh();
                return numFilesAndFoldersCompleted;
            }
        }

        /**When a user clicks the farkiv image, they are sent to the farkiv website
         */
        private void pictureBox1_Click(object sender, EventArgs e) {
            System.Diagnostics.Process.Start("https://farkiv.no");
        }

        private void OnChanged(object sender, FileSystemEventArgs e) {
            if (e.ChangeType != WatcherChangeTypes.Changed) {
                return;
            }
            try {
                if (File.GetAttributes(e.FullPath).HasFlag(FileAttributes.Directory)) {
                    FolderList.Remove(e.FullPath);
                    FolderList.Add(e.FullPath);
                }
                else
                    FileInCheck(e.FullPath);
            }
            catch (Exception ex) when 
                (ex is System.IO.DirectoryNotFoundException 
                || ex is System.IO.FileNotFoundException) {

                OnDeleted(sender, e);
            }
        }

        /**When a file or folder is created, it is added to the list and checked
         * Part of the watcher functions
         */
        private void OnCreated(object sender, FileSystemEventArgs e) {
            if (File.GetAttributes(e.FullPath).HasFlag(FileAttributes.Directory))
                FolderInLoop(e.FullPath);
            else if (e.FullPath.Substring(e.FullPath.LastIndexOf(@"\") + 1) != "Thumbs.db" &&
                        e.FullPath.Substring(e.FullPath.LastIndexOf(@"\") + 2) != "Thumbs.db" &&
                        e.FullPath.Substring(e.FullPath.LastIndexOf(@"\")) != "Thumbs.db")
                FileInCheck(e.FullPath);

            UpdateInScreen();
        }

        /**When a file or folder is removed, it is removed from the list
         * Part of the watcher functions
         */
        private void OnDeleted(object sender, FileSystemEventArgs e) {
            var fileType = e.FullPath.Substring(e.FullPath.LastIndexOf('.') + 1);
            if (fileType == e.FullPath) { 
                while (FolderList.Remove(e.FullPath));
                DeleteFilesInFolder(e.FullPath);
            }
            else
                switch (fileType) {
                    case "jpg":     while (JPGList.Remove(e.FullPath)); break;
                    case "jpeg":    while (JPGList.Remove(e.FullPath)); break;
                    case "jpe":     while (JPGList.Remove(e.FullPath)); break;
                    case "jif":     while (JPGList.Remove(e.FullPath)); break;
                    case "jfif":    while (JPGList.Remove(e.FullPath)); break;
                    case "jfi":     while (JPGList.Remove(e.FullPath)); break;
                    case "tif":     while (TIFList.Remove(e.FullPath)); break;
                    case "tiff":    while (TIFList.Remove(e.FullPath)); break;
                    case "gif":     while (GIFList.Remove(e.FullPath)); break;
                    case "png":     while (PNGList.Remove(e.FullPath)); break;
                    default:        while (MiscList.Remove(e.FullPath)); break;
                }

            UpdateInScreen();
        }

        private void DeleteFilesInFolder (string folderPath){
            for (int i = FolderList.Count - 1; i >= 0; i--) {
                if (FolderList[i].Contains(folderPath))
                    FolderList.RemoveAt(i);
            }
            for (int i = JPGList.Count - 1; i >= 0; i--) {
                if (JPGList[i].Contains(folderPath))
                    JPGList.RemoveAt(i);
            }
            for (int i = TIFList.Count - 1; i >= 0; i--) {
                if (TIFList[i].Contains(folderPath))
                    TIFList.RemoveAt(i);
            }
            for (int i = GIFList.Count - 1; i >= 0; i--) {
                if (GIFList[i].Contains(folderPath))
                    GIFList.RemoveAt(i);
            }
            for (int i = PNGList.Count - 1; i >= 0; i--) {
                if (PNGList[i].Contains(folderPath))
                    PNGList.RemoveAt(i);
            }
            for (int i = MiscList.Count - 1; i >= 0; i--) {
                if (MiscList[i].Contains(folderPath))
                    MiscList.RemoveAt(i);
            }
        }

        private static void OnRenamed(object sender, RenamedEventArgs e) {
            Console.WriteLine($"Renamed:");
            Console.WriteLine($"    Old: {e.OldFullPath}");
            Console.WriteLine($"    New: {e.FullPath}");
        }

        private static void OnError(object sender, ErrorEventArgs e) =>
            PrintException(e.GetException());

        private static void PrintException(Exception ex) {
            if (ex != null) {
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine("Stacktrace:");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine();
                PrintException(ex.InnerException);
            }
        }

        /**Goes through the folder, checking its contents and placing it in the FolderList
         * string folderPath = path of folder to check
         */
        private void FolderInLoop(string folderPath) {
            FolderList.Add(folderPath);
            if (Directory.GetDirectories(folderPath).Length != 0)
                foreach (string folder in Directory.GetDirectories(folderPath).ToList())
                    FolderInLoop(folder);
            if (Directory.GetFiles(folderPath).Length != 0)
                foreach (string file in Directory.GetFiles(folderPath).ToList())
                    FileInCheck(file);
        }

        /**Places the file in the correct list
         * string filePath = path of file to check
         */
        private void FileInCheck(string filePath) {
            switch (Path.GetExtension(filePath).ToLower()) {
                case ".jpg": JPGList.Add(filePath); break;
                case ".jpeg": JPGList.Add(filePath); break;
                case ".jpe": JPGList.Add(filePath); break;
                case ".jif": JPGList.Add(filePath); break;
                case ".jfif": JPGList.Add(filePath); break;
                case ".jfi": JPGList.Add(filePath); break;
                case ".tif": TIFList.Add(filePath); break;
                case ".tiff": TIFList.Add(filePath); break;
                case ".gif": GIFList.Add(filePath); break;
                case ".png": PNGList.Add(filePath); break;
                default: MiscList.Add(filePath); break;
            }
        }

        /**Updates the labels on the In screen with available values
         */
        private void UpdateInScreen() {
            string textToScreen = "";
            string numToScreen = "";
            if (JPGList.Any()) {
                textToScreen += "JPG:\n";
                numToScreen += JPGList.Count + "\n";
            }
            if (TIFList.Any()) {
                textToScreen += "TIF:\n";
                numToScreen += TIFList.Count + "\n";
            }
            if (GIFList.Any()) {
                textToScreen += "GIF:\n";
                numToScreen += GIFList.Count + "\n";
            }
            if (PNGList.Any()) {
                textToScreen += "PNG:\n";
                numToScreen += PNGList.Count + "\n";
            }
            if (MiscList.Any()) {
                textToScreen += "\nAndre filtyper:\n";
                numToScreen += "\n" + MiscList.Count + "\n";
            }
            if (FolderList.Any()) {
                textToScreen += "\nMapper:";
                numToScreen += "\n" + FolderList.Count;
            }
            if (textToScreen == "")
                textToScreen = "Ingen filer eller mapper";
            SetText(textToScreen, FilesAndFoldersBox);
            SetText(numToScreen, NumOfFilesAndFoldersBox);
        }

        private void SetText(string text, Label label) {
            if (label.InvokeRequired) {
                SetTextCallback s = new SetTextCallback(SetText);
                this.Invoke(s, new object[] { text, label });
            }
            else
                label.Text = text;
        }
    }
}
