using System;
using System.Collections.Generic;
using System.Text;

namespace MCiTunesSynchronizer
{
    public class PlaylistItem
    {
        public string Path = null;
        public bool Hide = false;
        public bool Rebuild = false;
        public bool Shuffle = false;
        public bool Selected = false;
        public bool Root = false;
        public bool Folder = false;
        public bool Smart = false;
        public bool Ticked = true;
        public bool RemoveTracks = false;
    }
}
