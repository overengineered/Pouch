using Newtonsoft.Json;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PouchShell
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFiles)]
    public class PouchContextMenu : SharpContextMenu
    {
        public class Item
        {
            public string Path { get; set; }
            public string SourceDir { get; set; }
        }

        public List<Item> loaded;

        private List<Item> GetConfig()
        {
            if (loaded == null)
            {
                var assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var json = File.ReadAllText(Path.Combine(assemblyDir, "pouch.json"));
                loaded = JsonConvert.DeserializeObject<List<Item>>(json);
            }
            return loaded;
        }

        protected override bool CanShowMenu()
        {
            List<string> selectionSet = SelectedItemPaths.ToList();
            if (selectionSet.Count != 1)
                return false;

            var result = GetConfig().Any(i => i.Path == selectionSet.First());
            return result;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            List<string> selectionSet = SelectedItemPaths.ToList();

            if (selectionSet.Count == 1)
            {
                return CreateSingleSelectionMenu(selectionSet[0]);
            }

            return null;
        }

        private ContextMenuStrip CreateSingleSelectionMenu(string path)
        {
            Item config = GetConfig().FirstOrDefault(i => i.Path == path);

            if (config == null)
                return null;

            ContextMenuStrip menu = new ContextMenuStrip();

            ToolStripMenuItem item = new ToolStripMenuItem
            {
                Text = Res.Replace_with,
                Image = Res.Pouch
            };

            DirectoryInfo directory = new DirectoryInfo(config.SourceDir);
            foreach(DirectoryInfo folder in directory.GetDirectories())
            {
                String pouched = Path.Combine(folder.FullName, Path.GetFileName(config.Path));
                if (!File.Exists(pouched))
                    continue;

                ToolStripMenuItem subitem = new ToolStripMenuItem
                {
                    Text = folder.Name,
                    Tag = new Tuple<String, String>(folder.FullName, config.Path)
                };
                subitem.Click += this.OnCopy;
                item.DropDownItems.Add(subitem);
            }

            menu.Items.Add(item);

            return menu;
        }

        private void OnCopy(object sender, EventArgs e)
        {
            ToolStripMenuItem item = (ToolStripMenuItem)sender;
            var operation = (Tuple<String, String>)item.Tag;

            // when using Windows Shell copy progress dialog, file overwrite confirmation is required :(
            // e.g. FileSystem.CopyFile(operation.Item1, operation.Item2, UIOption.AllDialogs, UICancelOption.DoNothing);
            // so instead copy via RoboCopy - at least some progress will be shown

            string targetDir = Path.GetDirectoryName(operation.Item2);
            string fileName = Path.GetFileName(operation.Item2);
            System.Diagnostics.Process.Start("robocopy.exe", String.Join(" ",
                ArgumentEscape(operation.Item1), ArgumentEscape(targetDir), ArgumentEscape(fileName)));
        }

        private static string ArgumentEscape(string arg)
        {
            if (arg.Contains(@" "))
                return @"""" + arg + @"""";
            return arg;
        }
    }
}
