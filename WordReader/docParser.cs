using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordReader
{
    class docParser
    {
        #region Classes
        /// <summary>
        /// OLE Compound File Binary class
        /// </summary>
        private class CompoundFileBinary
        {
            #region Structures
            /// <summary>
            /// Reserved special values
            /// </summary>
            private struct SpecialValues
            {
                /// <summary>
                /// Specifies a DIFAT sector in the FAT
                /// </summary>
                internal const uint DIFSECT = 0xFFFFFFFC;
                /// <summary>
                /// Specifies a FAT sector in the FAT
                /// </summary>
                internal const uint FATSECT = 0xFFFFFFFD;
                /// <summary>
                /// End of a linked chain of sectors
                /// </summary>
                internal const uint ENDOFCHAIN = 0xFFFFFFFE;
                /// <summary>
                /// Specifies an unallocated sector in the FAT, Mini FAT or DIFAT
                /// </summary>
                internal const uint FREESECT = 0xFFFFFFFF;
                /// <summary>
                /// Terminator or empty pointer if Directory Entry
                /// </summary>
                internal const uint NOSTREAM = 0xFFFFFFFF;
            }

            /// <summary>
            /// Compound File Header structure
            /// NOTE: for major version 3 CFHeader size is 512 bytes.
            /// NOTE: for major version 4 CFHeader size is 4096 bytes, so the remaining part (3584 bytes) if filled with zeros
            /// </summary>
            private struct CompoundFileHeader
            {
                /// <summary>
                /// Header Signature (MUST:  0xD0CF11E0A1B11AE1) [8 bytes]
                /// </summary>
                internal byte[] Signature;
                /// <summary>
                /// Header CLSID (MUST: all zeros) [16 bytes]
                /// </summary>
                internal byte[] CLSID;
                /// <summary>
                /// Minor Version (SHOULD: 0x3E00, if MajorVersion is 0x0300 or 0x0400) [2 bytes]
                /// </summary>
                internal byte[] MinorVersion;
                /// <summary>
                /// Major Version (MUST: 0x0300 (version 3) or 0x0400 (version 4)) [2 bytes]
                /// </summary>
                internal byte[] MajorVersion;
                /// <summary>
                /// Byte order (MUST: 0xFEFF) - little-endian [2 bytes]
                /// </summary>
                internal byte[] ByteOrder;
                /// <summary>
                /// Sector shift (MUST: 0x0009 (if major version is 3) or 0x000C (if major version is 4)) [2 bytes]
                /// </summary>
                internal byte[] SectorShift;
                /// <summary>
                /// Mini sector shift (sector size of the Mini Stream) (MUST: 0x0006) [2 bytes]
                /// </summary>
                internal byte[] MiniSectorShift;
                /// <summary>
                /// Reserved [6 bytes]
                /// </summary>
                internal byte[] Reserved;
                /// <summary>
                /// Number of Directory sectors (MUST: 0x0 if major version is 3) [1 uint = 4 bytes]
                /// </summary>
                internal uint NumDirSectors;
                /// <summary>
                /// Number of FAT sectors [1 uint = 4 bytes]
                /// </summary>
                internal uint NumFATSectors;
                /// <summary>
                /// First directory sector location - starting sector nmber for directory stream [1 uint = 4 bytes]
                /// </summary>
                internal uint FirstDirSectorLoc;
                /// <summary>
                /// Transaction signature number - how many times files was saved by implementation [1 uint = 4 bytes]
                /// </summary>
                internal uint TransSignNum;
                /// <summary>
                /// Max size of user-defined data stream (MUST: 0x00001000 = 4096) [1 uint = 4 bytes]
                /// </summary>
                internal uint MiniStreamCutoffSize;
                /// <summary>
                /// First mini FAT sector location - starting sector number for mini FAT [1 uint = 4 bytes]
                /// </summary>
                internal uint FirstMiniFATSectorLoc;
                /// <summary>
                /// Number of mini FAT sectors [1 uint = 4 bytes]
                /// </summary>
                internal uint NumMiniFATSectors;
                /// <summary>
                /// First DIFAT sector location - starting sector number for DIFAT [1 uint = 4 bytes]
                /// </summary>
                internal uint FirstDIFATSectorLoc;
                /// <summary>
                /// Number of DIFAT sectors [1 uint = 4 bytes]
                /// </summary>
                internal uint NumDIFATSectors;
                /// <summary>
                /// The first 109 FAT sector locations [109 uint = 436 bytes]
                /// </summary>
                internal uint[] DIFAT;
            }

            /// <summary>
            /// Compound File Directory Entry structure
            /// </summary>
            private struct DirectoryEntry
            {
                /// <summary>
                /// Directory Entry Name [64 bytes]
                /// </summary>
                internal string Name;
                /// <summary>
                /// Directory Entry Name in bytes (MUST: <=64) [2 bytes]
                /// </summary>
                internal uint NameLength;
                /// <summary>
                /// Object Type of the current directory entry
                /// (MUST: 0x00 (unknown or unallocated, 0x01 (Storage object), 0x02 (Stream object) OR 0x05 (Root Storage object)
                /// [1 byte]
                /// </summary>
                internal byte ObjectType;
                /// <summary>
                /// Color flag of the current directory entry (MUST: 0x00 (red), 0x01 (black)) [1 byte]
                /// </summary>
                internal byte ColorFlag;
                /// <summary>
                /// Left Sibling stream ID (MUST: 0xFFFFFFFF if there is no left sibling) [4 bytes]
                /// </summary>
                internal uint LeftSibling;
                /// <summary>
                /// Right Sibling stream ID (MUST: 0xFFFFFFFF if there is no right sibling) [4 bytes]
                /// </summary>
                internal uint RightSibling;
                /// <summary>
                /// Child object stream ID (MUST: 0xFFFFFFFF if there is no child objects) [4 bytes]
                /// </summary>
                internal uint Child;
                /// <summary>
                /// Object class GUID, if current entry is for a storage object or root storage object
                /// (MUST: all zeros for a stream object. MAY: all zeros for storage object or root storage object,
                /// thus indicating that no object class is associated with the storage)
                /// [16 bytes]
                /// </summary>
                internal byte[] CLSID;
                /// <summary>
                /// User-defined flags if current entry is for a storage object or root storage object
                /// (SHOULD: all zeros for a stream object) [4 bytes]
                /// </summary>
                internal byte[] StateBits;
                /// <summary>
                /// Creation Time for a storage object (MUST: all zeros for a stream object OR root storage object) [8 bytes]
                /// </summary>
                internal long CreationTime;
                /// <summary>
                /// Modification Time for a storage object (MUST: all zeros for a stream object. MAY: all zeros for a root storage object) [8 bytes]
                /// </summary>
                internal long ModifiedTime;
                /// <summary>
                /// Starting Sector Location  if this is a stream object (MUST: all zeros for a storage object.
                /// (MUST: first sector of the mini stream for a root storage object if the mini stream exists)
                /// [4 bytes]
                /// </summary>
                internal uint StartSectorLoc;
                /// <summary>
                /// Size of the user-defined data if this is a stream object. Size of the mini stream for a root storage object
                /// (MUST: all zeros for a storage object) [8 bytes]
                /// </summary>
                internal ulong StreamSizeV4;
                /// <summary>
                /// NOTE: THIS FIELD IS NOT IN REAL COMPOUND FILE DIRECTORY ENTRY STRUCTURE! I ADDED IT JUST FOR MY OWN CONVENIENCE!
                /// Same as StreamSizeV4, but used for version 3 compound files. That is StreamSizeV4 without most significant 32 bits.
                /// </summary>
                internal uint StreamSizeV3;
            }

            /// <summary>
            /// Folder Tree Entry structur (need to show folder structure of compound file as tree)
            /// </summary>
            private struct FolderTreeEntry
            {
                /// <summary>
                /// Level in folder tree of the current entry (Root Entry has level zero, every descent by Child ref. adds 1 to the level
                /// </summary>
                internal int TreeLevel;
                /// <summary>
                /// Name of the current entry
                /// </summary>
                internal string Name;
                /// <summary>
                /// Name of the parent entry
                /// </summary>
                internal string Parent;
            }
            #endregion

            #region Fields
            #region private
            private CompoundFileHeader CFHeader;    //Compound file header
            private uint[] DIFAT;                   //entire DIFAT array (from header + from DIFAT sectors)
            private uint[] FAT;                     //FAT array
            private uint[] miniFAT;                 //miniFAT array (from standart chain from header and FAT)
            private DirectoryEntry[] DEArray;       //Directory Entry Array
            private BinaryReader fileReader;        //doc binary reader
            #endregion

            #region protected internal
            /// <summary>
            /// True if Compound file header is OK
            /// </summary>
            protected internal bool CFHeaderIsOK;
            #endregion
            #endregion

            #region Constructors
            /// <summary>
            /// Class constructor
            /// </summary>
            /// <param name="reader">Binary reader for the Compound File</param>
            protected internal CompoundFileBinary(BinaryReader reader)
            {
                fileReader = reader;            //stored reader to the field
                CFHeaderIsOK = readCFHeader();  //read and checked the CF Header
                readDIFAT();                    //read DIFAT
                readFAT();                      //read FAT
                readminiFAT();                  //read miniFAT
                readDEArray();                  //read Directory Entry array
            }
            #endregion

            #region Methods
            #region private            
            /// <summary>
            /// Outputs Compound File Header to the console
            /// </summary>
            private void showFCHeader()
            {
                //writedown all bytes of the header to one byte-array
                int byteNumber = 0;
                byte[] Output = new byte[512];
                Array.Copy(CFHeader.Signature, 0, Output, 0, 8);
                Array.Copy(CFHeader.CLSID, 0, Output, (byteNumber += 8), 16);
                Array.Copy(CFHeader.MinorVersion, 0, Output, (byteNumber += 16), 2);
                Array.Copy(CFHeader.MajorVersion, 0, Output, (byteNumber += 2), 2);
                Array.Copy(CFHeader.ByteOrder, 0, Output, (byteNumber += 2), 2);
                Array.Copy(CFHeader.SectorShift, 0, Output, (byteNumber += 2), 2);
                Array.Copy(CFHeader.MiniSectorShift, 0, Output, (byteNumber += 2), 2);
                Array.Copy(CFHeader.Reserved, 0, Output, (byteNumber += 2), 6);
                Array.Copy(BitConverter.GetBytes(CFHeader.NumDirSectors), 0, Output, (byteNumber += 6), 4);
                Array.Copy(BitConverter.GetBytes(CFHeader.NumFATSectors), 0, Output, (byteNumber += 4), 4);
                Array.Copy(BitConverter.GetBytes(CFHeader.FirstDirSectorLoc), 0, Output, (byteNumber += 4), 4);
                Array.Copy(BitConverter.GetBytes(CFHeader.TransSignNum), 0, Output, (byteNumber += 4), 4);
                Array.Copy(BitConverter.GetBytes(CFHeader.MiniStreamCutoffSize), 0, Output, (byteNumber += 4), 4);
                Array.Copy(BitConverter.GetBytes(CFHeader.FirstMiniFATSectorLoc), 0, Output, (byteNumber += 4), 4);
                Array.Copy(BitConverter.GetBytes(CFHeader.NumMiniFATSectors), 0, Output, (byteNumber += 4), 4);
                Array.Copy(BitConverter.GetBytes(CFHeader.FirstDIFATSectorLoc), 0, Output, (byteNumber += 4), 4);
                Array.Copy(BitConverter.GetBytes(CFHeader.NumDIFATSectors), 0, Output, (byteNumber += 4), 4);
                foreach (uint difat in CFHeader.DIFAT)
                    Array.Copy(BitConverter.GetBytes(difat), 0, Output, (byteNumber += 4), 4);

                //output byte-array to the console
                showBytesInHEX(Output, "Compound file header", "end of header");
            }

            /// <summary>
            /// Outputs byte-array to console as HEX values
            /// </summary>
            /// <param name="Output">Byte-array to output</param>
            /// <param name="title">Title string before the output</param>
            /// <param name="ending">Ending string after the output</param>
            private void showBytesInHEX(byte[] Output, string title = "", string ending = "")
            {
                int byteNumber = 0;                     //number of the outputted bytes
                Console.WriteLine("\t" + title);        //output of the title
                Console.Write($"{byteNumber:X6}: ");    //number of the first byte in the current row
                foreach (byte o in Output)              //going through all the bytes
                {
                    if (byteNumber != 0 && (byteNumber % 16) == 0)  //when outputted 16 bytes
                    {
                        Console.WriteLine();                        //start the new row
                        Console.Write($"{byteNumber:X6}: ");        //and in that row output the number of the first byte in a row
                    }
                    Console.Write($"{o:X2}");                       //outputting the current byte
                    byteNumber++;                                   //increase number of the outputted bytes by one
                    if (byteNumber % 2 == 0) Console.Write(" ");    //output space between every two bytes
                }

                Console.WriteLine("\n\t" + ending);     //output of the ending string
            }

            /// <summary>
            /// Reads Compound File Header from the fileReader and checks it for errors
            /// </summary>
            /// <returns>TRUE if no errors found, FALSE otherwise</returns>
            private bool readCFHeader()
            {
                //rewind to the beginning of the file
                fileReader.BaseStream.Seek(0, SeekOrigin.Begin);

                //reading the header
                CFHeader.Signature = fileReader.ReadBytes(8);
                CFHeader.CLSID = fileReader.ReadBytes(16);
                CFHeader.MinorVersion = fileReader.ReadBytes(2);
                CFHeader.MajorVersion = fileReader.ReadBytes(2);
                CFHeader.ByteOrder = fileReader.ReadBytes(2);
                CFHeader.SectorShift = fileReader.ReadBytes(2);
                CFHeader.MiniSectorShift = fileReader.ReadBytes(2);
                CFHeader.Reserved = fileReader.ReadBytes(6);
                CFHeader.NumDirSectors = fileReader.ReadUInt32();
                CFHeader.NumFATSectors = fileReader.ReadUInt32();
                CFHeader.FirstDirSectorLoc = fileReader.ReadUInt32();
                CFHeader.TransSignNum = fileReader.ReadUInt32();
                CFHeader.MiniStreamCutoffSize = fileReader.ReadUInt32();
                CFHeader.FirstMiniFATSectorLoc = fileReader.ReadUInt32();
                CFHeader.NumMiniFATSectors = fileReader.ReadUInt32();
                CFHeader.FirstDIFATSectorLoc = fileReader.ReadUInt32();
                CFHeader.NumDIFATSectors = fileReader.ReadUInt32();

                CFHeader.DIFAT = new uint[109];
                for (int i = 0; i < 109; i++) CFHeader.DIFAT[i] = fileReader.ReadUInt32();
                                
                //standart (MUST) values of the fields in the Compound file header
                byte[] signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
                byte[] minorVersion = { 0x3E, 0x00 };
                byte[][] majorVersion = { new byte[] { 0x03, 0x00 }, new byte[] { 0x04, 0x00 } };
                byte[] byteOrder = { 0xFE, 0xFF };
                byte[][] sectorShift = { new byte[] { 0x09, 0x00 }, new byte[] { 0x0C, 0x00 } };
                byte[] miniSectorShift = { 0x06, 0x00 };

                //checking for errors (if any error found returning FALSE)
                if (!CFHeader.Signature.SequenceEqual(signature)) return false;
                if (!CFHeader.CLSID.SequenceEqual(new byte[16])) return false;
                if (!CFHeader.MinorVersion.SequenceEqual(minorVersion)) return false;
                if (!CFHeader.MajorVersion.SequenceEqual(majorVersion[0]) &&
                    !CFHeader.MajorVersion.SequenceEqual(majorVersion[1])) return false;
                if (!CFHeader.ByteOrder.SequenceEqual(byteOrder)) return false;
                if (!((CFHeader.MajorVersion.SequenceEqual(majorVersion[0]) &&
                    CFHeader.SectorShift.SequenceEqual(sectorShift[0])) ||
                    CFHeader.MajorVersion.SequenceEqual(majorVersion[1]) &&
                    CFHeader.SectorShift.SequenceEqual(sectorShift[1]))) return false;
                if (!CFHeader.MiniSectorShift.SequenceEqual(miniSectorShift)) return false;
                if (!(CFHeader.NumDirSectors == 0 && CFHeader.MajorVersion.SequenceEqual(majorVersion[0]))) return false;
                if (!(CFHeader.MiniStreamCutoffSize == 4096)) return false;

                //no errors found - returning TRUE
                return true;
            }

            /// <summary>
            /// Reads full DIFAT array from fileReader (from Header and from DIFAT sectors)
            /// </summary>
            private void readDIFAT()
            {
                //take DIFAT from CFHeader
                for (int i = 0; i < CFHeader.DIFAT.Length; i++)
                {
                    if (CFHeader.DIFAT[i] != SpecialValues.FREESECT)    //taking until found FREESECT value
                    {
                        //allocate memory
                        if (DIFAT == null) DIFAT = new uint[1];
                        else Array.Resize(ref DIFAT, DIFAT.Length + 1);
                        //copy the current value
                        DIFAT[i] = CFHeader.DIFAT[i];
                    }
                    else break;                                         //found FREESECT value - stop taking from CFHeader
                }

                //if there are no DIFAT sectors in the file, then we took the full DIFAT - returning from method
                if (CFHeader.NumDIFATSectors == 0 || CFHeader.FirstDIFATSectorLoc == SpecialValues.ENDOFCHAIN) return;

                //searching for DIFAT sectors in file and taking data from them
                uint numOfDIFATSectors = CFHeader.NumDIFATSectors;                                      //number of DIFAT sectors in the file
                uint curDIFATSEctorLoc = CFHeader.FirstDIFATSectorLoc;                                  //location of the current sector
                uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));    //size of one sector in the file
                int numEntriesInDIFAT = (int)(sectorSize - 4) / 4;                                      //number of DIFAT entries in one sector

                //while there still are unread DIFAT sectors in the file AND while haven't reached end of DIFAT sectors chain
                while (numOfDIFATSectors > 0 && curDIFATSEctorLoc != SpecialValues.ENDOFCHAIN)
                {
                    uint sectorOffset = (curDIFATSEctorLoc + 1) * sectorSize;   //offset of the current DIFAT sector in the file
                    fileReader.BaseStream.Seek(sectorOffset, SeekOrigin.Begin); //sought the offset in the file
                    for (int i = 0; i < numEntriesInDIFAT; i++)                 //reading all DIFAT entries from the current sector except the last one
                    {
                        uint tmp = fileReader.ReadUInt32(); //read one entry
                        if (tmp != SpecialValues.FREESECT)  //if it isn't empty
                        {
                            //reallocate memory
                            Array.Resize(ref DIFAT, DIFAT.Length + 1);
                            //save what we've read
                            DIFAT[DIFAT.Length - 1] = tmp;
                        }
                    }
                    curDIFATSEctorLoc = fileReader.ReadUInt32();                //reading the number of the next DIFAT sector in the chain
                    numOfDIFATSectors--;                                        //current sector reading is finished - decreasing the counter of DIFAT sectors
                }
            }

            /// <summary>
            /// Read FAT from the file
            /// </summary>
            private void readFAT()
            {
                uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));    //sector size in the file
                int numEntriesInFAT = (int)(sectorSize) / 4;                                            //number of FAT entries in one FAT sector

                //allocating memory
                FAT = new uint[CFHeader.NumFATSectors * numEntriesInFAT];

                for (int i = 0; i < DIFAT.Length; i++)  //going through the DIFAT table
                {
                    uint sectorOffset = (DIFAT[i] + 1) * sectorSize;            //offset of the current FAT sector in the file
                    fileReader.BaseStream.Seek(sectorOffset, SeekOrigin.Begin); //sought the offset
                    
                    //reading data from the file
                    for (int j = i * numEntriesInFAT; j < (i + 1) * numEntriesInFAT; j++)
                        FAT[j] = fileReader.ReadUInt32();
                }
            }

            /// <summary>
            /// Read the full miniFAT table from the file
            /// </summary>
            private void readminiFAT()
            {
                if (CFHeader.NumMiniFATSectors == 0)    //if there are no miniFAT sectors in the file
                {
                    miniFAT = null;
                    return;
                }

                uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));            //sector size in the file
                uint miniFATSectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.MiniSectorShift, 0)); //sector size in miniStream
                uint mfEntriesPerSector = sectorSize / 4;                                                       //number of miniFAT entries in one file sector
                uint numMiniFATEntries = CFHeader.NumMiniFATSectors * mfEntriesPerSector;                       //entire number of miniFAT entries in the file
                miniFAT = new uint[numMiniFATEntries];                                                          //allocating memory for miniFAT table
                uint currentminiFATsector = CFHeader.FirstMiniFATSectorLoc;                                     //number of the current miniFAT sector in the file
                int posInMiniFAT = 0;                                                                           //current position in the miniFAT array

                while (currentminiFATsector != SpecialValues.ENDOFCHAIN)                                        //while the end of FAT sectors containing miniFAT is not reached
                {
                    uint fileOffset = (currentminiFATsector + 1) * sectorSize;  //offset of the current sector in the file
                    fileReader.BaseStream.Seek(fileOffset, SeekOrigin.Begin);   //seek to the offset
                    byte[] readSector = fileReader.ReadBytes((int)sectorSize);  //read the current sector from the file
                    MemoryStream ms = new MemoryStream(readSector);             //created MemoryStream for the read sector
                    BinaryReader br = new BinaryReader(ms);                     //using the BinaryReader for the newly created MemoryStream

                    for (int i = 0; i < mfEntriesPerSector; i++)
                        miniFAT[posInMiniFAT + i] = br.ReadUInt32();            //reading miniFAT entries from the MemoryStream

                    posInMiniFAT += (int)mfEntriesPerSector;                    //increased the current position in the miniFAT array
                    currentminiFATsector = FAT[currentminiFATsector];           //moving to the next sector (with miniFAT) in the file
                }
            }

            /// <summary>
            /// Read Directory Entry array from the file
            /// </summary>
            private void readDEArray()
            {
                uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));                    //sector size in the file
                int numDirEntries = (int)(sectorSize) / 128;                                                            //number of Directory Entries in one file sector:
                                                                                                                        //one Directory Entry size equals 128 bytes, SO
                                                                                                                        //there are 4 Directory Entries in version 3 file OR
                                                                                                                        //there are 32 Directory Entries in version 4 file
                uint currentDirSector = CFHeader.FirstDirSectorLoc;                                                     //current sector with Directory Entries in the file
                int curDirSectorOrder = 0;                                                                              //counter of sector read from the file
                while (currentDirSector != SpecialValues.ENDOFCHAIN)                                                    //while end of stream not reached
                {
                    //allocating memory
                    if (DEArray == null) DEArray = new DirectoryEntry[numDirEntries];   //for the first sector
                    else Array.Resize(ref DEArray, DEArray.Length + numDirEntries);     //reallocating for the others sectors

                    uint sectorOffset = (currentDirSector + 1) * sectorSize;            //offset of the current sector in the file
                    fileReader.BaseStream.Seek(sectorOffset, SeekOrigin.Begin);         //seek to the offset

                    //moving through all directory entries in the current sector
                    for (int i = curDirSectorOrder * numDirEntries; i < (curDirSectorOrder + 1) * numDirEntries; i++)
                    {
                        //reading data of the current Directory Entry
                        DEArray[i].Name = Encoding.Unicode.GetString(fileReader.ReadBytes(64));
                        DEArray[i].NameLength = fileReader.ReadUInt16();
                        DEArray[i].ObjectType = fileReader.ReadByte();
                        DEArray[i].ColorFlag = fileReader.ReadByte();
                        DEArray[i].LeftSibling = fileReader.ReadUInt32();
                        DEArray[i].RightSibling = fileReader.ReadUInt32();
                        DEArray[i].Child = fileReader.ReadUInt32();
                        DEArray[i].CLSID = fileReader.ReadBytes(16);
                        DEArray[i].StateBits = fileReader.ReadBytes(4);
                        DEArray[i].CreationTime = fileReader.ReadInt64();
                        DEArray[i].ModifiedTime = fileReader.ReadInt64();
                        DEArray[i].StartSectorLoc = fileReader.ReadUInt32();
                        DEArray[i].StreamSizeV4 = fileReader.ReadUInt64();

                        //calculating StreamSizeV3
                        DEArray[i].StreamSizeV3 = Convert.ToUInt32(DEArray[i].StreamSizeV4);
                    }

                    currentDirSector = FAT[currentDirSector];   //get the number of the next sector in chain
                    curDirSectorOrder++;                        //increased the counter of read sectors
                }
            }

            /// <summary>
            /// Build the Folder Tree to later output it to the console
            /// </summary>            
            /// <param name="FTE">Array where the Folder Tree will be built</param>
            /// <param name="Id">Number of the current Directory Entry (NOTE: used in recursion, leave it empty)</param>
            private void buildFolderTree(ref FolderTreeEntry[] FTE, uint Id = 0)
            {
                //return if current Directiry Entry is NOSTREAM
                if (Id == SpecialValues.NOSTREAM) return;
                
                if (Id == 0)    //if we're in the Root Entry now
                {
                    FTE = new FolderTreeEntry[2];               //allocate memory for the first two entries in FolderTree entry array
                    FTE[0].TreeLevel = 0;                       //set the tree level of the Root Entry to zero

                    //filling the name of the current entry
                    FTE[0].Name = DEArray[0].Name.Substring(0, DEArray[0].Name.IndexOf('\0'));
                    FTE[0].Name += " <root storage>";

                    FTE[1].Parent = FTE[0].Name;                //filling the parent name for the child
                    FTE[1].TreeLevel = 1;                       //setting the tree level for the child to one

                    buildFolderTree(ref FTE, DEArray[0].Child); //go to the child   <- first entrance to the recursion

                    return;                                     //after filling the whole FolderTree entry array we'll come back here and will return from the method
                }

                //filling the current FolderTree Entry
                int curFTE = FTE.Length - 1;
                FTE[curFTE].Name = DEArray[Id].Name.Substring(0, DEArray[Id].Name.IndexOf('\0'));
                FTE[curFTE].Name += (DEArray[Id].ObjectType == 0x00) ? " <unknown>" : "";
                FTE[curFTE].Name += (DEArray[Id].ObjectType == 0x01) ? " <storage>" : "";
                FTE[curFTE].Name += (DEArray[Id].ObjectType == 0x02) ? " <stream>" : "";

                if (DEArray[Id].Child != SpecialValues.NOSTREAM)    //if current Directory Entry has a Child
                {
                    Array.Resize(ref FTE, FTE.Length + 1);                      //reallocate memory for FolderTree entry array
                    FTE[FTE.Length - 1].TreeLevel = FTE[curFTE].TreeLevel + 1;  //count and fill the tree level for the child
                    FTE[FTE.Length - 1].Parent = FTE[curFTE].Name;              //fill the parent name for the child
                    buildFolderTree(ref FTE, DEArray[Id].Child);                //go to the Child
                }

                if (DEArray[Id].LeftSibling != SpecialValues.NOSTREAM)  //if current Directory Entry has a Left Sibling
                {
                    Array.Resize(ref FTE, FTE.Length + 1);                  //reallocate memory for FolderTree entry array
                    FTE[FTE.Length - 1].TreeLevel = FTE[curFTE].TreeLevel;  //fill the tree level for the left sibling
                    FTE[FTE.Length - 1].Parent = FTE[curFTE].Parent;        //fill the parent name for the left sibling
                    buildFolderTree(ref FTE, DEArray[Id].LeftSibling);      //go to the left sibling
                }

                if (DEArray[Id].RightSibling != SpecialValues.NOSTREAM) //if current Directory Entry has a Right Sibling
                {
                    Array.Resize(ref FTE, FTE.Length + 1);                  //reallocate memory for FolderTree entry array
                    FTE[FTE.Length - 1].TreeLevel = FTE[curFTE].TreeLevel;  //fill the tree level for the right sibling
                    FTE[FTE.Length - 1].Parent = FTE[curFTE].Parent;        //fill the parent name for the right sibling
                    buildFolderTree(ref FTE, DEArray[Id].RightSibling);     //go to the right sibling
                }
            }

            /// <summary>
            /// Search for the stream in the file by the name specified
            /// </summary>            
            /// <param name="Name">Name of the stream to look for in the file</param>
            /// <param name="Paths">Here will be array of the Paths of the found streams (or null, if nothing is found)</param>
            /// <param name="StreamIds">Here will be array of the StreamIDs of the found streams (or null, if nothing is found)</param>
            /// <param name="Id">Number of the current Directory Entry (NOTE: used in recursion, leave it empty)</param>
            /// <param name="curPath">Path that is composed in this iteration (NOTE: used in recursion, leave it empty)</param>
            private void findInDEArray(string Name, ref string[] Paths, ref uint[] StreamIds, uint Id = 0, string curPath = "")
            {
                //return if current Directory Entry is NOSTREAM
                if (Id == SpecialValues.NOSTREAM) return;

                //checking the current Directory Entry
                string curName = DEArray[Id].Name.Substring(0, DEArray[Id].Name.IndexOf('\0'));
                if (curName.CompareTo(Name) == 0)   //if the name in the current Directory Entry is what we're looking for
                {
                    //---== saving StreamId
                    if (StreamIds == null) StreamIds = new uint[1];         //allocating memory if this is the first found stream
                    else Array.Resize(ref StreamIds, StreamIds.Length + 1); //reallocating if this stream is not the first found
                    StreamIds[StreamIds.Length - 1] = Id;                   //saving the current StreamId

                    //---== saving Path
                    if (Paths == null) Paths = new string[1];               //allocating memory if this is the first found stream
                    else Array.Resize(ref Paths, Paths.Length + 1);         //reallocating if this stream is not the first found
                    Paths[Paths.Length - 1] = curPath + "\\" + curName;     //saving the current Path
                }

                //go to Child
                string newPath = curPath;
                newPath += (newPath == "") ? "" : "\\";
                newPath += curName;
                findInDEArray(Name, ref Paths, ref StreamIds, DEArray[Id].Child, newPath);

                //go to Left Sibling
                findInDEArray(Name, ref Paths, ref StreamIds, DEArray[Id].LeftSibling, curPath);

                //go to Right Sibling
                findInDEArray(Name, ref Paths, ref StreamIds, DEArray[Id].RightSibling, curPath);
            }

            /// <summary>
            /// Search StreamID for the specified path
            /// </summary>            
            /// <param name="Path">Path to search for</param>
            /// <param name="foundId">Here will be the found StreamID (NOTE: always use returned bool value to determine whether the search was successful)</param>
            /// <param name="Id">Number of the current Directory Entry (NOTE: used in recursion, leave it empty)</param>
            /// <returns>TRUE is the search is successful (ID is stored in foundId). FALSE if no success (zero is stored in foundId).</returns>
            private bool findPathId(string Path, out uint foundId, uint Id = 0)
            {
                if (Id == SpecialValues.NOSTREAM)   //if current Directory Entry is NOSTREAM
                {
                    foundId = 0;    //return zero in foundId
                    return false;   //return FALSE
                }

                //getting the name of the current Directory Entry
                string curName = DEArray[Id].Name.Substring(0, DEArray[Id].Name.IndexOf('\0'));
                curName = curName.ToUpper();

                //getting the current path name (the name before the first \ in the search path)
                //(that will be the name of the storage or the name of the stream if we're already in the needed storage)
                string curPath = "";                                        //current path name
                int pathPos = Path.IndexOf('\\');                           //number of the first appearance of the '\' in the path
                if (pathPos != -1) curPath = Path.Substring(0, pathPos);    //if there is '\' in the path, we'll take the first storage name from the path
                else curPath = Path;                                        //if there is no '\' in the path, we'll take the stream name from the path

                //checking the current Directory Entry
                string nextPath = "";                                                   //next path name (Path without curPath part)
                if (curName.CompareTo(curPath) == 0)                                    //if current Directory Entry name equals current path name
                {
                    if (pathPos == -1)                                                  //AND if there was not '\' in the Path (which means that current path name is the search path)
                    {
                        foundId = Id;                                                   //return number of the current Directory Entry in foundId
                        return true;                                                    //return true
                    }
                    else                                                                //if there was '\' in the Path (which means that we maybe found the storage of the object that we're looking for)
                    {
                        nextPath = Path.Substring(pathPos + 1);                         //get the next path name
                        return findPathId(nextPath, out foundId, DEArray[Id].Child);    //and go to Child of current Directory Entry
                    }
                }

                //nothing found by now - going further
                nextPath = Path;    //we'll again use full path name
                //getting the number of the next Directory Entry (using the Red-Black Tree approach)
                uint nextId = 0;                                                                            //number of the next Directory Entry
                if (curPath.Length < DEArray[Id].NameLength / 2) nextId = DEArray[Id].LeftSibling;          //if current path name is shorter than current Directory Entry name, we'll go to the Left
                else if (curPath.Length > DEArray[Id].NameLength / 2) nextId = DEArray[Id].RightSibling;    //if current path name is longer than current Directory Entry name, we'll go to the Right
                                                                                                            //if lengths of those names are equal:
                else if (curPath.CompareTo(curName) < 0) nextId = DEArray[Id].LeftSibling;                  //if current path name is smaller (in alphabetic order) than current DE name, go to Left
                else nextId = DEArray[Id].RightSibling;                                                     //if neither of conditions worked (which means that current path name is bigger (in
                                                                                                            //alphabetic order), we'll go to Right

                //after that we just going to the direction that we've chosen
                return findPathId(nextPath, out foundId, nextId);
            }
            #endregion

            #region protected internal
            /// <summary>
            /// Outputs folder structure of Compound Binary File in console
            /// </summary>
            protected internal void showFolderTree()
            {
                FolderTreeEntry[] FTE = null;                           //Folder Tree Entry array

                buildFolderTree(ref FTE);                               //building FolderTree entry array

                //---== outputting folder tree to console using FolderTree entry array
                //output the header with the filename to the console
                FileStream fs = (FileStream)fileReader.BaseStream;
                Console.Clear();
                Console.WriteLine("Folder Tree of " + fs.Name);
                Console.WriteLine();

                //output folder tree
                Console.WriteLine(FTE[0].Name);                             //name of the first entry
                for (int i = 1; i < FTE.Length; i++)                        //moving through the other entries
                {
                    //--= drawing pseudographics
                    //we'll look at the tree levels first
                    for (int l = 1; l < FTE[i].TreeLevel; l++)              //moving throught all possible tree levels that are smaller than level of the current entry
                    {
                        int j = 0;
                        for (j = i + 1; j < FTE.Length; j++)                //moving through all remained entries in array
                        {
                            if (FTE[j].TreeLevel == l)                      //if we found an entry with current smaller level
                            {
                                Console.Write("│ ");                        //we must draw vertical line (because one of the parents of the current entry has childs among the remaining entries)
                                break;                                      //and break the current cicle (because one child is enough)
                            }
                        }
                        if (j == FTE.Length) Console.Write("  ");           //if we found no entry with current smaller level, we'll draw empty space (one of the parents of the current entry
                                                                            //has no children among the remaining entries)
                    }
                    //after looking at tree levels we must look at siblings
                    bool hasSiblingsFurther = false;
                    for (int j = i + 1; j < FTE.Length; j++)                //moving through all the remaining entries
                        if (FTE[j].Parent.CompareTo(FTE[i].Parent) == 0)    //if any of the remaining entries has such parent as the current entry does
                        {
                            hasSiblingsFurther = true;                      //we'll say that the current entry has siblings further
                            break;                                          //and yes, one sibling is enough
                        }
                    if (hasSiblingsFurther) Console.Write("├─");            //if we found a sibling, we'll draw the cross (because parent of the current entry has childs among remaining entries)
                    else Console.Write("└─");                               //and if no siblings was found, we'll draw corner (because current entry is the last child among the outputted ones)

                    Console.WriteLine(FTE[i].Name);                         //after all we can output the name of the current entry
                }

                Console.WriteLine();
            }

            /// <summary>
            /// Search for the stream in the file by the name specified
            /// </summary>
            /// <param name="Name">Name of the stream to look for in the file</param>
            /// <param name="Paths">Here will be array of the Paths of the found streams (or null, if nothing is found)</param>
            /// <param name="StreamIds">Here will be array of the StreamIDs of the found streams (or null, if nothing is found)</param>
            /// <returns>true if found at least one stream, false if nothing was found</returns>
            protected internal bool findStream(string Name, ref string[] Paths, ref uint[] StreamIds)
            {
                findInDEArray(Name, ref Paths, ref StreamIds);  //searchig for the stream

                if (Paths == null || StreamIds == null)         //if nothing was found
                {
                    Paths = null;           //write null to Paths
                    StreamIds = null;       //write null to StreamIds
                    return false;           //return false
                }

                if (Paths.Length != StreamIds.Length)           //if search result is incorrect (Paths and StreamIds must have the same number of items)
                {
                    Paths = null;           //write null to Paths
                    StreamIds = null;       //write null to StreamIds
                    return false;           //return false
                }

                return true;                                    //return true if search is successful
            }

            /// <summary>
            /// Search for the stream in the file by the name specified
            /// </summary>
            /// <param name="Name">Name of the stream to look for in the file</param>
            /// <param name="Paths">Here will be array of the Paths of the found streams (or null, if nothing is found)</param>
            /// <returns>true if found at least one stream, false if nothing was found</returns>
            protected internal bool findStream(string Name, ref string[] Paths)
            {
                //we'll simply use the overloaded version: findStream(string,ref string[],ref uint[])
                uint[] StreamIds = null;
                return findStream(Name, ref Paths, ref StreamIds);
            }

            /// <summary>
            /// Get StreamID for the path specified
            /// </summary>
            /// <param name="Path">Path to look for</param>
            /// <param name="Id">Here will be the found StreamID (NOTE: always use returned bool value to determine whether the search was successful)</param>
            /// <returns>TRUE is the search is successful (ID is stored in foundId). FALSE if no success (zero is stored in foundId).</returns>
            protected internal bool getPathId(string Path, out uint Id)
            {
                return findPathId(Path.ToUpper(), out Id);  //simply using the analogous private method
            }

            /// <summary>
            /// Get stream from the file by its StreamID
            /// </summary>
            /// <param name="Id">ID of the stream</param>
            /// <returns>MemoryStream object which represents the needed stream (null if no stream with such ID in the file or the object with specified StreamID is storage, root storage or unallocated)</returns>
            protected internal MemoryStream getStream(uint Id)
            {
                byte[] byteStream = null;   //stream as a byte array
                MemoryStream memStream;     //stream as a MemoryStream
                
                switch (DEArray[Id].ObjectType) //everything depends on the type of the object with specified ID
                {
                    case 0x00:                  //OBJECT TYPE: unknown OR unallocated
                        return null;                                                                                //we'll just return null
                    case 0x01:                  //OBJECT TYPE: storage
                        return null;                                                                                //we'll just return null
                    case 0x02:                  //OBJECT TYPE: stream
                                                                                                                    //calculate the size of the specified stream
                        ulong streamSize = (CFHeader.MajorVersion.SequenceEqual(new byte[] { 0x04, 0x00 })) ?       //if file is of version 4
                            DEArray[Id].StreamSizeV4 :                                                              //we will use StreamSizeV4 as the size of the stream
                            DEArray[Id].StreamSizeV3;                                                               //otherwise (version 3) we'll use StreamSizeV3
                        byteStream = new byte[streamSize];                                                          //allocate memory for the stream

                        //First of all we must understand which stream we need to read from the file:
                        //  - if size of the stream is greater or equal to the cutoff for the mini-stream, we will read our stream from sectors referenced with FAT;
                        //  - but if it is smaller, we will read the mini-stream from the file and after that we will read our stream from mini-stream sectors referenced with miniFAT.

                        uint readId = (streamSize >= CFHeader.MiniStreamCutoffSize) ?                               //if size of the stream is greater or equal to the cutoff for mini-stream
                            Id :                                                                                    //we'll read specified stream right away
                            0;                                                                                      //otherwise we'll begin with reading the mini-stream (Id=0 in DE-array)

                        uint curSector = DEArray[readId].StartSectorLoc;                                            //get the first sector to read
                        uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));        //sector size in the file

                                                                                                                    //calculate the size of the stream which we'll read from the file
                                                                                                                    //(it'll be specified stteam or mini-stream)
                        ulong fileBufferSize = (CFHeader.MajorVersion.SequenceEqual(new byte[] { 0x04, 0x00 })) ?   //if file is of version 4
                            DEArray[readId].StreamSizeV4 :                                                          //we will use StreamSizeV4 as the size
                            DEArray[readId].StreamSizeV3;                                                           //otherwise (version 3) we'll use StreamSizeV3
                        byte[] fileBuffer = new byte[fileBufferSize];                                               //buffer for the data read from the file
                        byte[] readSector = new byte[sectorSize];                                                   //current read from the file sector
                        int posInArray = 0;                                                                         //our current position in fileBuffer
                        
                        //reading from the file
                        while (curSector != SpecialValues.ENDOFCHAIN)                                               //while the end of the sector chain isn't reached
                        {
                            uint fileOffset = (curSector + 1) * sectorSize;                                         //offset of the current sector in the file
                            fileReader.BaseStream.Seek(fileOffset, SeekOrigin.Begin);                               //seek to the current offset
                            readSector = fileReader.ReadBytes((int)sectorSize);                                     //read the current sector from the file
                            curSector = FAT[curSector];                                                             //found the next sector
                            if (curSector == SpecialValues.ENDOFCHAIN)                                              //if current sector was the last one in the chain (next sector is ENDOFCHAIN)
                            {
                                int numLastBytes = (int)Math.IEEERemainder(fileBufferSize, sectorSize);             //calculate number of bytes used for the stream in the last sector
                                readSector.Take(numLastBytes).ToArray().CopyTo(fileBuffer, posInArray);             //copy that bytes to the fileBuffer
                            }
                            else                                                                                    //if current sector wasn't the last one
                            {
                                readSector.CopyTo(fileBuffer, posInArray);                                          //copy whole sector to the fileBuffer
                                posInArray += (int)sectorSize;                                                      //increase the counter of current position in fileBuffer

                            }
                        }
                        //finished reading from the file

                        if (streamSize >= CFHeader.MiniStreamCutoffSize)                                            //if size of the specified stream is greater or equal to the cutoff for mini-stream
                        {
                                                                                                                    //then we've read right what was specified
                            fileBuffer.CopyTo(byteStream, 0);                                                       //copy the read data to the byte array for the stream
                        }
                        else                                                                                        //if size is not greater or equal
                        {
                                                                                                                    //then we've read the mini-stream
                                                                                                                    //and now we will read the specified stream from the mini-stream
                            MemoryStream miniStream = new MemoryStream(fileBuffer);                                 //created MemoryStream from the fileBuffer
                            BinaryReader brMiniStream = new BinaryReader(miniStream);                               //created BinaryReader for the new MemoryStream

                            curSector = DEArray[Id].StartSectorLoc;                                                 //first sector in mini-stream with the stream specified
                            sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.MiniSectorShift, 0));     //sector size in the mini-stream
                            Array.Resize(ref readSector, (int)sectorSize);                                          //reallocated memory for the current read from mini-stream sector
                            posInArray = 0;                                                                         //now it is the current position in the byteStream
                            
                            //reading from the mini-stream
                            while (curSector != SpecialValues.ENDOFCHAIN)                                           //while the end of the sector chain isn't reached
                            {
                                uint msOffset = (curSector + 1) * sectorSize;                                       //offset of the current sector in the mini-stream
                                miniStream.Seek(msOffset, SeekOrigin.Begin);                                        //seek to the current offset
                                readSector = brMiniStream.ReadBytes((int)sectorSize);                               //read the current sector from the mini-stream
                                curSector = miniFAT[curSector];                                                     //found the next sector
                                if (curSector == SpecialValues.ENDOFCHAIN)                                          //if current sector was the last one in the chain (next sector is ENDOFCHAIN)
                                {
                                    int numLastBytes = (int)Math.IEEERemainder(streamSize, sectorSize);             //calculate number of bytes used for the stream in the last sector
                                    readSector.Take(numLastBytes).ToArray().CopyTo(byteStream, posInArray);         //copy that bytes to the byteStream
                                }
                                else                                                                                //if current sector wasn't the last one
                                {
                                    readSector.CopyTo(byteStream, posInArray);                                      //copy whole sector to the byteStream
                                    posInArray += (int)sectorSize;                                                  //increase the counter of current position in byteStream

                                }
                            }
                        }
                        break;
                    case 0x05:                  //OBJECT TYPE: root storage                        
                        return null;                                                                                //we'll just return null
                    default:                    //OBJECT TYPE UNKNOWN
                        return null;                                                                                //we'll just return null
                }

                memStream = new MemoryStream(byteStream);                                                           //create MemoryStream from the byteStream
                return memStream;                                                                                   //return MemoryStream
            }

            /// <summary>
            /// Get stream from the file by its Path
            /// </summary>
            /// <param name="Path">Path of the stream</param>
            /// <returns>MemoryStream object which represents the needed stream (null if no stream with such Path in the file or the object with specified Path is storage, root storage or unallocated)</returns>
            protected internal MemoryStream getStream(string Path)
            {
                //we will simply use the overloaded version of this method: getStream(uint)
                uint Id;                                            //Id of the stream specified with the Path
                if (getPathId(Path, out Id)) return getStream(Id);  //if specified stream is found, we'll return it's MemoryStream
                return null;                                        //we will return null if specified stream was not found
            }

            /// <summary>
            /// Close all reader streams and clear all class fields
            /// </summary>
            protected internal void closeReader()
            {
                //close binaryreader and filereader stream
                fileReader.Close();

                //clear all fields' values
                CFHeaderIsOK = false;
                DIFAT = null;
                FAT = null;
                miniFAT = null;
                DEArray = null;
                CFHeader.ByteOrder = null;
                CFHeader.CLSID = null;
                CFHeader.DIFAT = null;
                CFHeader.FirstDIFATSectorLoc = 0;
                CFHeader.FirstDirSectorLoc = 0;
                CFHeader.FirstMiniFATSectorLoc = 0;
                CFHeader.MajorVersion = null;
                CFHeader.MiniSectorShift = null;
                CFHeader.MiniStreamCutoffSize = 0;
                CFHeader.MinorVersion = null;
                CFHeader.NumDIFATSectors = 0;
                CFHeader.NumDirSectors = 0;
                CFHeader.NumFATSectors = 0;
                CFHeader.NumMiniFATSectors = 0;
                CFHeader.Reserved = null;
                CFHeader.SectorShift = null;
                CFHeader.Signature = null;
                CFHeader.TransSignNum = 0;                
            }
            #endregion
            #endregion
        }

        /// <summary>
        /// Mapped to Unicode values for FcCompressed structure
        /// </summary>
        private static class MappedToUnicode
        {
            internal static Dictionary<byte, char> values = new Dictionary<byte, char>();    //Dictionary with mapped values
            static MappedToUnicode()
            {
                //fill in the mapped values to the dictionary
                values.Add(0x82, (char)0x201A);
                values.Add(0x83, (char)0x0192);
                values.Add(0x84, (char)0x201E);
                values.Add(0x85, (char)0x2026);
                values.Add(0x86, (char)0x2020);
                values.Add(0x87, (char)0x2021);
                values.Add(0x88, (char)0x02C6);
                values.Add(0x89, (char)0x2030);
                values.Add(0x8A, (char)0x0160);
                values.Add(0x8B, (char)0x2039);
                values.Add(0x8C, (char)0x0152);
                values.Add(0x91, (char)0x2018);
                values.Add(0x92, (char)0x2019);
                values.Add(0x93, (char)0x201C);
                values.Add(0x94, (char)0x201D);
                values.Add(0x95, (char)0x2022);
                values.Add(0x96, (char)0x2013);
                values.Add(0x97, (char)0x2014);
                values.Add(0x98, (char)0x02DC);
                values.Add(0x99, (char)0x2122);
                values.Add(0x9A, (char)0x0161);
                values.Add(0x9B, (char)0x203A);
                values.Add(0x9C, (char)0x0153);
                values.Add(0x9F, (char)0x0178);
            }
        }

        /// <summary>
        /// Character & paragraph property modifiers
        /// </summary>
        private static class PropSPRM
        {
            #region enums
            /// <summary>
            /// Property visibility flags
            /// </summary>
            private enum PropVisFlag
            {
                /// <summary>
                /// Character is visible
                /// </summary>
                Visible,
                /// <summary>
                /// Character is invisible
                /// </summary>
                InVisible,
                /// <summary>
                /// Character visibility depends on the character
                /// </summary>
                CharDep
            }

            /// <summary>
            /// Character visibility flags
            /// </summary>
            private enum CharVisFlag
            {
                /// <summary>
                /// Character is visible
                /// </summary>
                Visible,
                /// <summary>
                /// Character is invisible
                /// </summary>
                InVisible,
                /// <summary>
                /// Character opens invisible region of characters
                /// </summary>
                OpenRgn,
                /// <summary>
                /// Character closes invisible region of characters
                /// </summary>
                CloseRgn,
                /// <summary>
                /// Character visibility depends on paragraph properties operand
                /// </summary>
                Paragraph
            }
            #endregion

            #region fields
            /// <summary>
            /// If true the character is currently inside the region of invisible text
            /// </summary>
            private static bool InvisibleRgn = false;

            /// <summary>
            /// If true the character is visible
            /// </summary>
            private static bool Visible = true;

            /// <summary>
            /// Stores information about visibility of characters with actual properties
            /// </summary>
            private static Dictionary<ushort, PropVisFlag> PropVisibility = new Dictionary<ushort, PropVisFlag>();

            /// <summary>
            /// Stores information about visibility of some characters
            /// </summary>
            private static Dictionary<char, CharVisFlag> CharVisibility = new Dictionary<char, CharVisFlag>();
            #endregion

            #region constructors
            /// <summary>
            /// Class constructor
            /// </summary>
            static PropSPRM()
            {
                //---== fill PropVisibility dictionary
                //Character properties
                PropVisibility.Add(0x0800, PropVisFlag.Visible);     //sprmCFRMarkDel
                PropVisibility.Add(0x0801, PropVisFlag.Visible);     //sprmCFRMarkIns
                PropVisibility.Add(0x0802, PropVisFlag.InVisible);   //sprmCFFldVanish
                PropVisibility.Add(0x6A03, PropVisFlag.CharDep);    //sprmCPicLocation
                PropVisibility.Add(0x4804, PropVisFlag.Visible);     //sprmCIbstRMark
                PropVisibility.Add(0x6805, PropVisFlag.Visible);     //sprmCDttmRMark
                PropVisibility.Add(0x0806, PropVisFlag.InVisible);   //sprmCFData
                PropVisibility.Add(0x4807, PropVisFlag.Visible);     //sprmCIdslRMark
                PropVisibility.Add(0x6A09, PropVisFlag.Visible);     //sprmCSymbol
                PropVisibility.Add(0x080A, PropVisFlag.InVisible);    //sprmCFOle2
                PropVisibility.Add(0x2A0C, PropVisFlag.Visible);     //sprmCHighlight
                PropVisibility.Add(0x0811, PropVisFlag.Visible);     //sprmCFWebHidden
                PropVisibility.Add(0x6815, PropVisFlag.Visible);     //sprmCRsidProp
                PropVisibility.Add(0x6816, PropVisFlag.Visible);     //sprmCRsidText
                PropVisibility.Add(0x6817, PropVisFlag.Visible);     //sprmCRsidRMDel
                PropVisibility.Add(0x0818, PropVisFlag.InVisible);   //sprmCFSpecVanish
                PropVisibility.Add(0xC81A, PropVisFlag.Visible);     //sprmCFMathPr
                PropVisibility.Add(0x4A30, PropVisFlag.Visible);     //sprmCIstd
                PropVisibility.Add(0xCA31, PropVisFlag.Visible);     //sprmCIstdPermute
                PropVisibility.Add(0x2A33, PropVisFlag.Visible);     //sprmCPlain
                PropVisibility.Add(0x2A34, PropVisFlag.Visible);     //sprmCKcd
                PropVisibility.Add(0x0835, PropVisFlag.Visible);     //sprmCFBold
                PropVisibility.Add(0x0836, PropVisFlag.Visible);     //sprmCFItalic
                PropVisibility.Add(0x0837, PropVisFlag.Visible);     //sprmCFStrike
                PropVisibility.Add(0x0838, PropVisFlag.Visible);     //sprmCFOutline
                PropVisibility.Add(0x0839, PropVisFlag.Visible);     //sprmCFShadow
                PropVisibility.Add(0x083A, PropVisFlag.Visible);     //sprmCFSmallCaps
                PropVisibility.Add(0x083B, PropVisFlag.Visible);     //sprmCFCaps
                PropVisibility.Add(0x083C, PropVisFlag.InVisible);   //sprmCFVanish
                PropVisibility.Add(0x2A3E, PropVisFlag.Visible);     //sprmCKul
                PropVisibility.Add(0x8840, PropVisFlag.Visible);     //sprmCDxaSpace
                PropVisibility.Add(0x2A42, PropVisFlag.Visible);     //sprmCIco
                PropVisibility.Add(0x4A43, PropVisFlag.Visible);     //sprmCHps
                PropVisibility.Add(0x4845, PropVisFlag.Visible);     //sprmCHpsPos
                PropVisibility.Add(0xCA47, PropVisFlag.Visible);     //sprmCMajority
                PropVisibility.Add(0x2A47, PropVisFlag.Visible);     //sprmCIss
                PropVisibility.Add(0x484B, PropVisFlag.Visible);     //sprmCHpsKern
                PropVisibility.Add(0x484E, PropVisFlag.Visible);     //sprmCHresi
                PropVisibility.Add(0x4A4F, PropVisFlag.Visible);     //sprmCRgFtc0
                PropVisibility.Add(0x4A50, PropVisFlag.Visible);     //sprmCRgFtc1
                PropVisibility.Add(0x4A51, PropVisFlag.Visible);     //sprmCRgFtc2
                PropVisibility.Add(0x4852, PropVisFlag.Visible);     //sprmCCharScale
                PropVisibility.Add(0x2A53, PropVisFlag.Visible);     //sprmCFDStrike
                PropVisibility.Add(0x0854, PropVisFlag.Visible);     //sprmCFImprint
                PropVisibility.Add(0x0855, PropVisFlag.CharDep);   //sprmCFSpec
                PropVisibility.Add(0x0856, PropVisFlag.InVisible);    //sprmCFObj
                PropVisibility.Add(0xCA57, PropVisFlag.Visible);     //sprmCPropRMark90
                PropVisibility.Add(0x0858, PropVisFlag.Visible);     //sprmCFEmboss
                PropVisibility.Add(0x2859, PropVisFlag.Visible);     //sprmCSfxText
                PropVisibility.Add(0x085A, PropVisFlag.Visible);     //sprmCFBiDi
                PropVisibility.Add(0x085C, PropVisFlag.Visible);     //sprmCFBoldBi
                PropVisibility.Add(0x085D, PropVisFlag.Visible);     //sprmCFItalicBi
                PropVisibility.Add(0x4A5E, PropVisFlag.Visible);     //sprmCFtcBi
                PropVisibility.Add(0x485F, PropVisFlag.Visible);     //sprmCLidBi
                PropVisibility.Add(0x4A60, PropVisFlag.Visible);     //sprmCIcoBi
                PropVisibility.Add(0x4A61, PropVisFlag.Visible);     //sprmCHpsBi
                PropVisibility.Add(0xCA62, PropVisFlag.Visible);     //sprmCDispFldRMark
                PropVisibility.Add(0x4863, PropVisFlag.Visible);     //sprmCIbstRMarkDel
                PropVisibility.Add(0x6864, PropVisFlag.Visible);     //sprmCDttmRMarkDel
                PropVisibility.Add(0x6865, PropVisFlag.Visible);     //sprmCBrc80
                PropVisibility.Add(0x4866, PropVisFlag.Visible);     //sprmCShd80
                PropVisibility.Add(0x4867, PropVisFlag.Visible);     //sprmCIdslRMarkDel
                PropVisibility.Add(0x0868, PropVisFlag.Visible);     //sprmCFUsePgsuSettings
                PropVisibility.Add(0x486D, PropVisFlag.Visible);     //sprmCRgLid0_80
                PropVisibility.Add(0x486E, PropVisFlag.Visible);     //sprmCRgLid1_80
                PropVisibility.Add(0x286F, PropVisFlag.Visible);     //sprmCIdctHint
                PropVisibility.Add(0x6870, PropVisFlag.Visible);     //sprmCCv
                PropVisibility.Add(0xCA71, PropVisFlag.Visible);     //sprmCShd
                PropVisibility.Add(0xCA72, PropVisFlag.Visible);     //sprmCBrc
                PropVisibility.Add(0x4873, PropVisFlag.Visible);     //sprmCRgLid0
                PropVisibility.Add(0x4874, PropVisFlag.Visible);     //sprmCRgLid1
                PropVisibility.Add(0x0875, PropVisFlag.Visible);     //sprmCFNoProof
                PropVisibility.Add(0xCA76, PropVisFlag.Visible);     //sprmCFitText
                PropVisibility.Add(0x6877, PropVisFlag.Visible);     //sprmCCvUl
                PropVisibility.Add(0xCA78, PropVisFlag.Visible);     //sprmCFELayout
                PropVisibility.Add(0x2879, PropVisFlag.Visible);     //sprmCLbcCRJ
                PropVisibility.Add(0x0882, PropVisFlag.Visible);     //sprmCFComplexScripts
                PropVisibility.Add(0x2A83, PropVisFlag.Visible);     //sprmCWall
                PropVisibility.Add(0xCA85, PropVisFlag.Visible);     //sprmCCnf
                PropVisibility.Add(0x2A86, PropVisFlag.Visible);     //sprmCNeedFontFixup
                PropVisibility.Add(0x6887, PropVisFlag.Visible);     //sprmCPbiIBullet
                PropVisibility.Add(0x4888, PropVisFlag.Visible);     //sprmCPbiGrf
                PropVisibility.Add(0xCA89, PropVisFlag.Visible);     //sprmCPropRMark
                PropVisibility.Add(0x2A90, PropVisFlag.InVisible);   //sprmCFSdtVanish

                //Paragraph properties
                PropVisibility.Add(0x4600, PropVisFlag.Visible);    //sprmPIstd
                PropVisibility.Add(0xC601, PropVisFlag.Visible);    //sprmPIstdPermute
                PropVisibility.Add(0x2602, PropVisFlag.Visible);    //sprmPIncLvl
                PropVisibility.Add(0x2403, PropVisFlag.Visible);    //sprmPJc80
                PropVisibility.Add(0x2405, PropVisFlag.Visible);    //sprmPFKeep
                PropVisibility.Add(0x2406, PropVisFlag.Visible);    //sprmPFKeepFollow
                PropVisibility.Add(0x2407, PropVisFlag.Visible);    //sprmPFPageBreakBefore
                PropVisibility.Add(0x260A, PropVisFlag.Visible);    //sprmPIlvl
                PropVisibility.Add(0x460B, PropVisFlag.Visible);    //sprmPIlfo
                PropVisibility.Add(0x240C, PropVisFlag.Visible);    //sprmPFNoLineNumb
                PropVisibility.Add(0xC60D, PropVisFlag.Visible);    //sprmPChgTabsPapx
                PropVisibility.Add(0x840E, PropVisFlag.Visible);    //sprmPDxaRight80
                PropVisibility.Add(0x840F, PropVisFlag.Visible);    //sprmPDxaLeft80
                PropVisibility.Add(0x4610, PropVisFlag.Visible);    //sprmPNest80
                PropVisibility.Add(0x8411, PropVisFlag.Visible);    //sprmPDxaLeft180
                PropVisibility.Add(0x6412, PropVisFlag.Visible);    //sprmPDyaLine
                PropVisibility.Add(0xA413, PropVisFlag.Visible);    //sprmPDyaBefore
                PropVisibility.Add(0xA414, PropVisFlag.Visible);    //sprmPDyaAfter
                PropVisibility.Add(0xC615, PropVisFlag.Visible);    //sprmPChgTabs
                PropVisibility.Add(0x2416, PropVisFlag.Visible);    //sprmPFInTable
                PropVisibility.Add(0x2417, PropVisFlag.CharDep);    //sprmPFTtp
                PropVisibility.Add(0x8418, PropVisFlag.Visible);    //sprmPDxaAbs
                PropVisibility.Add(0x8419, PropVisFlag.Visible);    //sprmPDyaAbs
                PropVisibility.Add(0x841A, PropVisFlag.Visible);    //sprmPDxaWidth
                PropVisibility.Add(0x261B, PropVisFlag.Visible);    //sprmPPc
                PropVisibility.Add(0x2423, PropVisFlag.Visible);    //sprmPWr
                PropVisibility.Add(0x6424, PropVisFlag.Visible);    //sprmPBrcTop80
                PropVisibility.Add(0x6425, PropVisFlag.Visible);    //sprmPBrcLeft80
                PropVisibility.Add(0x6426, PropVisFlag.Visible);    //sprmPBrcBottom80
                PropVisibility.Add(0x6427, PropVisFlag.Visible);    //sprmPBrcRight80
                PropVisibility.Add(0x6428, PropVisFlag.Visible);    //sprmPBrcBetween80
                PropVisibility.Add(0x6629, PropVisFlag.Visible);    //sprmPBrcBar80
                PropVisibility.Add(0x242A, PropVisFlag.Visible);    //sprmPFNoAutoHyph
                PropVisibility.Add(0x442B, PropVisFlag.Visible);    //sprmPWHeightAbs
                PropVisibility.Add(0x442C, PropVisFlag.Visible);    //sprmPDcs
                PropVisibility.Add(0x442D, PropVisFlag.Visible);    //sprmPShd80
                PropVisibility.Add(0x842E, PropVisFlag.Visible);    //sprmPDyaFromText
                PropVisibility.Add(0x842F, PropVisFlag.Visible);    //sprmPDxaFromText
                PropVisibility.Add(0x2430, PropVisFlag.Visible);    //sprmPFLocked
                PropVisibility.Add(0x2431, PropVisFlag.Visible);    //sprmPFWidowControl
                PropVisibility.Add(0x2433, PropVisFlag.Visible);    //sprmPFKinsoku
                PropVisibility.Add(0x2434, PropVisFlag.Visible);    //sprmPFWordWrap
                PropVisibility.Add(0x2435, PropVisFlag.Visible);    //sprmPFOverflowPunct
                PropVisibility.Add(0x2436, PropVisFlag.Visible);    //sprmPFTopLinePunct
                PropVisibility.Add(0x2437, PropVisFlag.Visible);    //sprmPFAutoSpaceDE
                PropVisibility.Add(0x2438, PropVisFlag.Visible);    //sprmPFAutoSpaceDN
                PropVisibility.Add(0x4439, PropVisFlag.Visible);    //sprmPWAlignFont
                PropVisibility.Add(0x443A, PropVisFlag.Visible);    //sprmPFrameTextFlow
                PropVisibility.Add(0x2640, PropVisFlag.Visible);    //sprmPOutLvl
                PropVisibility.Add(0x2441, PropVisFlag.Visible);    //sprmPFBiDi
                PropVisibility.Add(0x2443, PropVisFlag.Visible);    //sprmPFNumRMIns
                PropVisibility.Add(0xC645, PropVisFlag.Visible);    //sprmPNumRM
                PropVisibility.Add(0x6646, PropVisFlag.Visible);    //sprmPHugePapx
                PropVisibility.Add(0x2447, PropVisFlag.Visible);    //sprmPFUsePgsuSettings
                PropVisibility.Add(0x2448, PropVisFlag.Visible);    //sprmPFAdjustRight
                PropVisibility.Add(0x6649, PropVisFlag.Visible);    //sprmPItap
                PropVisibility.Add(0x664A, PropVisFlag.Visible);    //sprmPDtap
                PropVisibility.Add(0x244B, PropVisFlag.CharDep);    //sprmPFInnerTableCell
                PropVisibility.Add(0x244C, PropVisFlag.Visible);    //sprmPFInnerTtp
                PropVisibility.Add(0xC64D, PropVisFlag.Visible);    //sprmPShd
                PropVisibility.Add(0xC64E, PropVisFlag.Visible);    //sprmPBrcTop
                PropVisibility.Add(0xC64F, PropVisFlag.Visible);    //sprmPBrcLeft
                PropVisibility.Add(0xC650, PropVisFlag.Visible);    //sprmPBrcBottom
                PropVisibility.Add(0xC651, PropVisFlag.Visible);    //sprmPBrcRight
                PropVisibility.Add(0xC652, PropVisFlag.Visible);    //sprmPBrcBetween
                PropVisibility.Add(0xC653, PropVisFlag.Visible);    //sprmPBrcBar
                PropVisibility.Add(0x4455, PropVisFlag.Visible);    //sprmPDxcRight
                PropVisibility.Add(0x4456, PropVisFlag.Visible);    //sprmPDxcLeft
                PropVisibility.Add(0x4457, PropVisFlag.Visible);    //sprmPDxcLeft1
                PropVisibility.Add(0x4458, PropVisFlag.Visible);    //sprmPDylBefore
                PropVisibility.Add(0x4459, PropVisFlag.Visible);    //sprmPDylAfter
                PropVisibility.Add(0x245A, PropVisFlag.Visible);    //sprmPFOpenTch
                PropVisibility.Add(0x245B, PropVisFlag.Visible);    //sprmPFDyaBeforeAuto
                PropVisibility.Add(0x245C, PropVisFlag.Visible);    //sprmPFDyaAfterAuto
                PropVisibility.Add(0x845D, PropVisFlag.Visible);    //sprmPDxaRight
                PropVisibility.Add(0x845E, PropVisFlag.Visible);    //sprmPDxaLeft
                PropVisibility.Add(0x465F, PropVisFlag.Visible);    //sprmPNest
                PropVisibility.Add(0x8460, PropVisFlag.Visible);    //sprmPDxaLeft1
                PropVisibility.Add(0x2461, PropVisFlag.Visible);    //sprmPJc
                PropVisibility.Add(0x2462, PropVisFlag.Visible);    //sprmPFNoAllowOverlap
                PropVisibility.Add(0x2664, PropVisFlag.Visible);    //sprmPWall
                PropVisibility.Add(0x6465, PropVisFlag.Visible);    //sprmPIpgp
                PropVisibility.Add(0xC666, PropVisFlag.Visible);    //sprmPCnf
                PropVisibility.Add(0x6467, PropVisFlag.Visible);    //sprmPRsid
                PropVisibility.Add(0xC669, PropVisFlag.Visible);    //sprmPIstdListPermute
                PropVisibility.Add(0x646B, PropVisFlag.Visible);    //sprmPTableProps
                PropVisibility.Add(0xC66C, PropVisFlag.Visible);    //sprmPTIstdInfo
                PropVisibility.Add(0x246D, PropVisFlag.Visible);    //sprmPFContextualSpacing
                PropVisibility.Add(0xC66F, PropVisFlag.Visible);    //sprmPPropRMark
                PropVisibility.Add(0x2470, PropVisFlag.Visible);    //sprmPFMirrorIndents
                PropVisibility.Add(0x2471, PropVisFlag.Visible);    //sprmPTtwo

                //---== fill CharVisibility Dictionary
                CharVisibility.Add('\u0001', CharVisFlag.InVisible);    //picture location
                CharVisibility.Add('\u0002', CharVisFlag.InVisible);    //auto-numbered footnote reference
                CharVisibility.Add('\u0003', CharVisFlag.InVisible);    //short horizontal line
                CharVisibility.Add('\u0004', CharVisFlag.InVisible);    //long horizontal line
                CharVisibility.Add('\u0005', CharVisFlag.InVisible);    //annotation reference character
                CharVisibility.Add('\u0008', CharVisFlag.InVisible);    //drawn object
                CharVisibility.Add('\u0013', CharVisFlag.OpenRgn);      //field begin character
                CharVisibility.Add('\u0014', CharVisFlag.CloseRgn);     //field separator character
                CharVisibility.Add('\u0015', CharVisFlag.InVisible);    //field end character
                CharVisibility.Add('\u0028', CharVisFlag.Visible);      //symbol '('
                CharVisibility.Add('\u003C', CharVisFlag.OpenRgn);      //start of a structured document tag bookmark range
                CharVisibility.Add('\u003E', CharVisFlag.CloseRgn);     //end of a structured document tag bookmark range
                CharVisibility.Add('\u2002', CharVisFlag.Visible);      //en space
                CharVisibility.Add('\u2003', CharVisFlag.Visible);      //em space
                CharVisibility.Add('\u0007', CharVisFlag.InVisible);    //cell mark
                CharVisibility.Add('\u000D', CharVisFlag.Paragraph);    //paragraph mark

            }
            #endregion

            #region methods
            /// <summary>
            /// Check visibility of the specified character with specified property
            /// </summary>
            /// <param name="pr">Property of the character</param>
            /// <param name="ch">Character</param>
            /// <returns>TRUE if character is visible</returns>
            internal static bool IsVisible(Prl[] prls, char ch)
            {
                //
                //NOTE: we will not forget that this class, all of its fields and methods are static!
                //      It saves its condition (values of fields). And we'll use this feature of static classes.
                //

                CharVisFlag cvf = CharVisFlag.Visible;                          //flag of character visibility according to the character itself
                bool hasCharInformation =                                       //flags successfullness of retrieving infomation in CharVisibility dictionary
                    CharVisibility.TryGetValue(ch, out cvf);                    //trying to find current character in CharVisibility dictionary

                switch (cvf)                                                    //switch current character visibility infomration
                {
                    case CharVisFlag.Visible: Visible = true; break;            //visible
                    case CharVisFlag.InVisible: Visible = false; break;         //invisible
                    default: Visible = true; break;                             //visible in any other cases
                }

                if (prls == null)                                               //if there are no prls for ch 
                    return Visible && !InvisibleRgn;                            //return true if character is visible and is not inside the invisible region

                foreach (Prl prl in prls)                                       //moving through all prl-s
                {
                    PropVisFlag pvf = PropVisFlag.Visible;                      //flag of character visibility according to its properties
                    if (PropVisibility.TryGetValue(prl.sprm.sprm, out pvf))     //trying to find current property in PropVisibility dictionary 
                                                                                //and to get information of visibility of the current character from it
                    {
                        switch(pvf)                                             //if property was found choose visibility flag
                        {
                            case PropVisFlag.Visible:                           //character with current property is visible - we'll do nothing
                                break;
                            case PropVisFlag.InVisible:                         //character with current property is invisible
                                Visible = false;                                //set visibility of character to false
                                break;
                            case PropVisFlag.CharDep:                           //visibility of the character with current property depends of the character itself
                                if (hasCharInformation)                         //if CharVisibility dictionary has information about current character
                                {
                                    switch (cvf)                                //choose visibility flag (not used previously cases)
                                    {
                                        case CharVisFlag.OpenRgn:               //current character opens the region of invisible characters
                                            InvisibleRgn = true;                //set the flag that says us that current character is inside invisible region to true
                                            Visible = false;                    //but current character is invisible anyway
                                            break;
                                        case CharVisFlag.CloseRgn:              //current character closes the region of invisible characters
                                            InvisibleRgn = false;               //set the flag that says us that current character is inside invisible region to false
                                            Visible = false;                    //but current character is invisible anyway
                                            break;
                                        case CharVisFlag.Paragraph:             //current character properties depends of the paragraph properties operand
                                            if (prl.sprm.sprm == 0x2417 &&      //current Sprm is sprmPFTtp
                                                ch == '\u0007' ||               //and current character is a cell mark
                                                prl.sprm.sprm == 0x244B &&      //OR current Sprm is sprmPFInnerTableCell
                                                ch == '\u000D')                 //and current character is a paragraph mark
                                            {
                                                Visible = false;                //set visibility of character to false
                                            }
                                            break;
                                    }
                                }
                                break;
                        }
                    }
                }
                return Visible && !InvisibleRgn;                                //return true if character is visible and is not inside the invisible region
            }
            #endregion
        }
        #endregion

        #region Structures
        /// <summary>
        /// Contains information about the document and specifies the file pointers to various portions that make up the document
        /// </summary>
        private struct FIB
        {
            //
            //NOTE: That is a partial structure from [MS-DOC] v20190319.
            //      I used only the fileds that are needed in this class
            //

            /// <summary>
            /// {WD.FIB.base} The FIBbase structure [off.: 0; len.: 32 bytes]
            /// </summary>
            internal FIBbase _base;

            /// <summary>
            /// {WD.FIB.fibRgLw} The FibRgLw97 structure [off.: 64; len.: 88 bytes]
            /// </summary>
            internal FibRgLw97 fibRgLw;

            /// <summary>
            /// {WD.FIB.fibRgFcLcbBlob} The FibRgFcLcb structure [off.: 154; len.: variable]
            /// </summary>
            internal FibRgFcLcb fibRgFcLcbBlob;

            /// <summary>
            /// Indicates whether Fib is clear or not
            /// NOTE: I added this field for my convenience!
            /// </summary>
            internal bool IsClear;
        }

        /// <summary>
        /// The FibBase structure
        /// </summary>
        private struct FIBbase
        {
            //
            //NOTE: That is a partial structure from [MS-DOC] v20190319.
            //      I used only the fileds that are needed in this class
            //

            /// <summary>
            /// {WD.FIB.FibBase.A-M} Bit-field that specifies a lot of stuff [off.: 10; len.: 2 bytes]
            /// </summary>
            internal byte[] AtoM;

            /// <summary>
            /// {WD.FIB.FibBase.fWhichTblStm (bit 6 (G) of AtoM)} Specifies the Table stream to which the FIB refers (true - 1Table, false - 0Table) [1 bit]
            /// </summary>
            internal bool fWhichTblStm;
        }

        /// <summary>
        /// The FibRgLw97 structure
        /// </summary>
        private struct FibRgLw97
        {
            //
            //NOTE: That is a partial structure from [MS-DOC] v20190319.
            //      I used only the fileds that are needed in this class
            //

            /// <summary>
            /// {WD.Fib.FibRgLw.ccpText} Count of CPs in the Main Document (MUST: >=0) [off.: 12;len.: 4 bytes]
            /// </summary>
            internal int ccpText;
        }

        /// <summary>
        /// The FibRgFcLcb97 structure
        /// </summary>
        private struct FibRgFcLcb97
        {
            //
            //NOTE: That is a partial structure from [MS-DOC] v20190319.
            //      I used only the fileds that are needed in this class
            //

            /// <summary>
            /// {Fib.FibRgFcLcb97.fcPlcfBteChpx} Offset of PlcBteChpx in the Table stream [off.: 96;len.: 4 bytes]
            /// </summary>
            internal uint fcPlcfBteChpx;

            /// <summary>
            /// {Fib.FibRgFcLcb97.lcbPlcfBteChpx} Size in bytes of PlcBteChpx in the Table Stream [off.: 100;len.: 4 bytes]
            /// </summary>
            internal uint lcbPlcfBteChpx;

            /// <summary>
            /// {Fib.FibRgFcLcb97.fcPlcfBtePapx} Offset of PlcBtePapx in the Table stream [off.: 104; len.: 4 bytes]
            /// </summary>
            internal uint fcPlcfBtePapx;

            /// <summary>
            /// {Fib.FibRgFcLcb97.lcbPlcfBtePapx} Size in bytes of PlcBtePapx in the Table stream [off.: 108; len.: 4 bytes]
            /// </summary>
            internal uint lcbPlcfBtePapx;

            /// <summary>
            /// {WD.Fib.fibRgFcLcb97.fcClx} Offset of the Clx in the Table stream [off.: 264;len.: 4 bytes]
            /// </summary>
            internal uint fcClx;

            /// <summary>
            /// {WD.Fib.fibRgFcLcb97.lcbClx} Size in bytes of the Clx in the Table stream (MUST >0) [off.: 268;len.: 4 bytes]
            /// </summary>
            internal uint lcbClx;

           
        }

        /// <summary>
        /// The FibRgFcLcb structure
        /// </summary>
        private struct FibRgFcLcb
        {
            //
            //NOTE: That is a partial structure from [MS-DOC] v20190319.
            //      I used only the fileds that are needed in this class
            //

            /// <summary>
            /// {WD.FIB.fibRgFcLcbBlob.fibRgFcLcb97} The FibRgFcLcb97 part of the variable length structure FibRgFcLcb
            /// </summary>
            internal FibRgFcLcb97 fibRgFcLcb97;
        }

        /// <summary>
        /// The Clx structure
        /// </summary>
        private struct CLX
        {
            /// <summary>
            /// An array of Prc structures
            /// </summary>
            internal Prc[] RgPrc;

            /// <summary>
            /// A Pcdt structure
            /// </summary>
            internal Pcdt pcdt;

            /// <summary>
            /// Indicates whether Clx is clear or not
            /// NOTE: I added this field for my convenience!
            /// </summary>
            internal bool IsClear;
        }

        /// <summary>
        /// The Pcdt structure
        /// </summary>
        private struct Pcdt
        {
            /// <summary>
            /// {Clx.Pcdt.clxt} (MUST: 0x02) [off.: 0; len.: 1 byte]
            /// </summary>
            internal byte clxt;

            /// <summary>
            /// {Clx.Pcdt.lcb} Specifies the size of the PlcPcd which follows [off.: 1; len.: 4 bytes]
            /// </summary>
            internal uint lcb;

            /// <summary>
            /// {Clx.Pcdt.PlcPcd} The PlcPcd structure [off.: 5; len.: lcb bytes]
            /// </summary>
            internal PlcPcd plcPcd;
        }

        /// <summary>
        /// The PlcPcd structure
        /// </summary>
        private struct PlcPcd
        {
            /// <summary>
            /// {Clx.Pcdt.PlcPcd.aCP} An array of CPs that specifies the starting points of text ranges [off.: 0; len.: variable]
            /// </summary>
            internal uint[] aCP;

            /// <summary>
            /// {Clx.Pcdt.PlcPcd.aPcd} An array of Pcds that specify the location of text in the WordDocument stream and any additional properties of text [off.: variable; len.: variable]
            /// </summary>
            internal Pcd[] aPcd;
        }

        /// <summary>
        /// Specifies the set of properties for document content that is referenced by a Pcd structure
        /// </summary>
        private struct Prc
        {
            /// <summary>
            /// (MUST: 0x01) [1 byte]
            /// </summary>
            internal byte clxt;

            /// <summary>
            /// A PrcData structure that specifies a set of properties
            /// </summary>
            internal PrcData data;
        }

        /// <summary>
        /// The PrcData structure
        /// </summary>
        private struct PrcData
        {
            /// <summary>
            /// {Clx.Prc.cbGrpprl} Size in bytes of the GrpPrl which follows (MUST be less or equal to 0x3FA2) [off.: 0;len.: 2 bytes]
            /// </summary>
            internal short cbGrpprl;

            /// <summary>
            /// {Clx.Prc.GrpPrl} An array of Prl structures [off.: 2; len.: cbGrpprl bytes]
            /// </summary>
            internal Prl[] GrpPrl;
        }

        /// <summary>
        /// Specifies the location of text in the WordDocument stream and additional properties for this text
        /// </summary>
        private struct Pcd
        {
            //
            //NOTE: That is a partial structure from [MS-DOC] v20190319.
            //      I used only the fileds that are needed in this class
            //

            /// <summary>
            /// An FcCompressed structure that specifies the location of the text in the WordDocument stream
            /// </summary>
            internal FcCompressed fc;

            /// <summary>
            /// Prm structure. Either Prm0 (if fComplex = false) or Prm1 (if fComplex = true) [2 bytes]
            /// </summary>
            internal Prm prm;            
        }

        /// <summary>
        /// Prm structure. Either Prm0 (if fComplex = false) or Prm1 (if fComplex = true) [2 bytes]
        /// </summary>
        private struct Prm
        {
            /// <summary>
            /// Bit that specifies what is the type of this Prm: 0 or 1 [1 bit]
            /// </summary>
            internal bool fComplex;

            /// <summary>
            /// Part of Prm0 structure. Specifies the Sprm to apply to the document [7 bits]
            /// </summary>
            internal byte Prm0_isprm;

            /// <summary>
            /// Part of Prm0 structure. Operand for the Sprm specified by isprm [1 byte]
            /// </summary>
            internal byte Prm0_val;

            /// <summary>
            /// Part of Prm1 structure. Zero-based index of a Prc in ClxRgPrc [15 bits]
            /// </summary>
            internal ushort Prm1_igrpprl;
        }

        /// <summary>
        /// Specifies the location of text in the WordDocument stream
        /// </summary>
        private struct FcCompressed
        {
            //
            //NOTE: That is a partial structure from [MS-DOC] v20190319.
            //      I used only the fileds that are needed in this class
            //

            /// <summary>
            /// Offset in WordDocument stream where the text starts [30 bits]
            /// If fCompressed=false, the text is an array of 16-bit Unicode characters starting at offset fc
            /// If fCompressed=true, the text is an array of 8-bit ANSI characters starting at offset fc/2, except mapped to Unicode values
            /// </summary>
            internal uint fc;
            /// <summary>
            /// A bit that specified whether the text is compressed [1 bit]
            /// </summary>
            internal bool fCompressed;
        }

        /// <summary>
        /// Specifies the location of the ChpxFkp structure if WordDocument stream
        /// </summary>
        private struct PnFkpChpx
        {
            /// <summary>
            /// Offset of the ChpxFkp structure in WordDocument stream. Offset = pn*512 [22 bits]
            /// </summary>
            internal uint pn;

            /// <summary>
            /// MUST be ignored [10 bits]
            /// </summary>
            internal ushort unused;
        }

        /// <summary>
        /// Maps text to its character properties
        /// </summary>
        private struct ChpxFkp
        {
            /// <summary>
            /// Offset in the WordDocument stream where a run of text begins [crun+1 elements each 4 bytes long]
            /// </summary>
            internal uint[] rgfc;

            /// <summary>
            /// Specifies the offset of one of the Chpxs whithin this ChpxFkp. Offset is computed by muliplying this value by 2
            /// (MUST: offset OR 0 which means that no Chpx for this item) [crun elements each 1 byte long]
            /// </summary>
            internal byte[] rgb;

            /// <summary>
            /// Array of Chpx structures
            /// </summary>
            internal Chpx[] chpx;

            /// <summary>
            /// Number of runs of text in this ChpxFkp (MUST: >= 0x01 AND NOT EXCEED 0x65) [last 1 byte of this ChpxFkp]
            /// </summary>
            internal byte crun;
        }

        /// <summary>
        /// Maps paragraphs, table rows, and table cells to their properties
        /// </summary>
        private struct PapxFkp
        {
            /// <summary>
            /// Offset in the WordDocument stream where a paragraph of text begins, or where an end of row mark exists [cpara+1 elements each 4 bytes long]
            /// </summary>
            internal uint[] rgfc;

            /// <summary>
            /// Specifies the offset of one of the PapxInFkp whithin this PapxFkp [cpara elements each 13 bytes long]
            /// </summary>
            internal BxPap[] rgbx;

            /// <summary>
            /// Array of PapxInFkp structures
            /// </summary>
            internal PapxInFkp[] papxInFkp;

            /// <summary>
            /// Number of runs of text in this ChpxFkp (MUST: >= 0x01 AND NOT EXCEED 0x1D) [last 1 byte of this PapxFkp]
            /// </summary>
            internal byte cpara;
        }

        /// <summary>
        /// Specifies the offset of a PapxInFkp in PapxFkp
        /// </summary>
        private struct BxPap
        {
            /// <summary>
            /// {PapxFkp.BxPap.bOffset} Specifies the offset of a PapxInFkp in a PapxFkp. The offset of the PapxInFkp is bOffset*2.
            /// If bOffset is 0 then there is no PapxInFkp for this paragraph/ [off.: 0; len.: 1 byte]
            /// </summary>
            internal byte bOffset;

            /// <summary>
            /// {PapxFkp.BxPap.reserved} Specifies version-specific paragraph height information (SHOULD: 0 and be ignored) [off.: 1; len.: 12 bytes]
            /// </summary>
            internal byte[] reserved;
        }

        /// <summary>
        /// Specifies a set of text properties
        /// </summary>
        private struct PapxInFkp
        {
            /// <summary>
            /// Size of the grpprlInPapx
            /// If this value is not 0, the grpprlInPapx is 2*cb-1 bytes long.
            /// If this value is 0, the size is specified by the first byte of grpprlInPapx.
            /// [off.: 0; len.: 1 byte]
            /// </summary>
            internal byte cb;

            /// <summary>
            /// Size of the grpprlInPapx if cb is zero
            /// The grpprlInPapx is 2*cb_ bytes long
            /// (MUST: >=1) [off.: 1; len.: 1 byte]
            /// </summary>
            internal byte cb_;

            /// <summary>
            /// The GrpPrlAndIstd structure
            /// </summary>
            internal GrpPrlAndIstd grpprlInPapx;
        }

        /// <summary>
        /// Specifies the style and properties that are applied to a paragraph, a table row, or a table cell
        /// </summary>
        private struct GrpPrlAndIstd
        {
            /// <summary>
            /// Specifies the style that is applied to this paragraph, cell marker or table row marker [off.: 0; len.: 2 bytes]
            /// </summary>
            internal short istd;

            /// <summary>
            /// Size of GrpPrl, in bytes (used only in case of sprmPHugePapx property applied (when PrcData structure replaces GrpPrlAndIstd))
            /// [off.: 0; len.: 2 bytes]
            /// </summary>
            internal short cbGrpprl;

            /// <summary>
            /// Specifies the properties of this paragraph, table row, or table cell [off.: 2; len.: variable]
            /// </summary>
            internal Prl[] grpprl;
        }

        /// <summary>
        /// Specifies the set of properties for text
        /// </summary>
        private struct Chpx
        {
            /// <summary>
            /// Specifies the size of grpprl in bytes [1 byte]
            /// </summary>
            internal byte cb;

            /// <summary>
            /// Specifies the properties
            /// </summary>
            internal Prl[] grpprl;
        }

        /// <summary>
        /// A Sprm that is followed by an operand
        /// </summary>
        private struct Prl
        {
            /// <summary>
            /// Specifies the property being modified [2 bytes]
            /// </summary>
            internal Sprm sprm;

            /// <summary>
            /// Operand for the sprm [variable length specified by sprm.spra]
            /// </summary>
            internal byte[] operand;
        }

        /// <summary>
        /// Specifies a modification to a property of a character, paragrahp, table or section
        /// </summary>
        private struct Sprm
        {
            /// <summary>
            /// In combination with fSpec specifies the property being modified [9 bits]
            /// </summary>
            internal ushort ispmd;  //Formula to calculate this field: ispmd = sprm & 0x01FF

            /// <summary>
            /// In combination with ispmd specifies the property being modified [1 bit]
            /// </summary>
            internal bool fSpec;    //Formula to calculate this field: fSpec = (sprm/512) & 0x0001

            /// <summary>
            /// Specifies the kind of document content to which this Sprm applies
            /// (MUST: 1 - modifies a paragraph property, 2 - character, 3 - picture, 4 - section, 5 - table property)
            /// [3 bits]
            /// </summary>
            internal byte sgc;      //Formula to calculate this field: sgc = (sprm/1024) & 0x0007

            /// <summary>
            /// Size of the operand of this Sprm
            /// (MUST: 0 - ToggleOperand (1 byte in size); 1 - 1 byte; 2, 4 or 5 - 2 bytes; 3 - 4 bytes; 7 - 3 bytes;
            /// 6 - operand is of variable length, the first byte of the operand indicates the size of the rest of the operand, except in the cases of sprmTDefTable and sprmPChgTabs
            /// [3 bits]
            /// </summary>
            internal byte spra;     //Formula to calculate this field: spra = sprm / 8192

            /// <summary>
            /// A 16-bit integer representation of the current Sprm [2 bytes]
            /// </summary>
            internal ushort sprm;
        }

        /// <summary>
        /// Maps the offsets of text in the WordDocument stream to the character properties of that text
        /// </summary>
        private struct PLCBTECHPX
        {
            /// <summary>
            /// {PlcBteChpx.aFC} Specifies an offset in the WordDocument stream where text begins [off.: 0; len.: variable]
            /// </summary>
            internal uint[] aFC;

            /// <summary>
            /// {PlcBteChps.aPnFkpChpx} An array of PnFkpChpx structures [off.: variable; len.: variable]
            /// </summary>
            internal PnFkpChpx[] aPnBteChpx;
        }

        /// <summary>
        /// Specifies paragraph, table row, or table cell properties
        /// </summary>
        private struct PLCBTEPAPX
        {
            /// <summary>
            /// {PlcBtePapx.aFC} Specifies an offset in the WordDocument stream where text begins [off.: 0; len.: variable]
            /// </summary>
            internal uint[] aFC;

            /// <summary>
            /// {PlcBtePapx.aPnBtePapx} An array of PnFkpPapx structures [off.: variable; len.: variable]
            /// </summary>
            internal PnFkpChpx[] aPnBtePapx;        //I used PnFkpChpx structures simply because PnFkpPapx and PnFkpChpx are absolutly identical
        }
        #endregion

        #region Fields
        #region private
        private CompoundFileBinary CFB = null;      //class for reading the Compound Binary File
        private MemoryStream WDStream = null;       //WordDocument stream (Main Document)
        private string WDStreamPath = null;         //Path to WordDocument stream if it is read from CFB (Main Document)
        private MemoryStream TableStream = null;    //Table stream (of the Main Document)
        private MemoryStream DataStream = null;     //Data stream (of the Main Document)
        private FIB Fib;                            //FIB in the WDStream
        private CLX Clx;                            //Clx in the TableStream
        private PLCBTECHPX PlcBteChpx;              //PlcBteChpx in the Table stream
        private ChpxFkp[] aChpxFkp;                 //array of ChpxFkp structures in the WDStream
        private PLCBTEPAPX PlcBtePapx;              //PlcBtePapx int the Table stream
        private PapxFkp[] aPapxFkp;                 //array of PapxFkp structures in the WDStream
        #endregion

        #region protected internal
        /// <summary>
        /// True is file exists, is OK and is a Word Binary File
        /// </summary>
        protected internal bool docIsOK = false;
        #endregion
        #endregion

        #region Constructors
        /// <summary>
        /// Class constructor
        /// </summary>
        /// <param name="filePath">Path to the file to be read</param>
        protected internal docParser(string filePath)
        {
            FileStream fileStream;                                                      //FileStream for the specified file
            try
            {
                fileStream = new FileStream(filePath, FileMode.Open);                   //trying to open specified file
            }
            catch(FileNotFoundException)                                                //if file not found
            {
                ConsoleColor cc = Console.ForegroundColor;                              //save current console foreground color
                Console.ForegroundColor = ConsoleColor.Red;                             //set color to RED
                Console.WriteLine("\nFile " + filePath + " not found!");                //output error message to console
                Console.ForegroundColor = cc;                                           //revert console to the saved foreground color
                return;                                                                 //and return from constructor
            }

                                                                                        //if file successfully opened
            BinaryReader fileReader = new BinaryReader(fileStream, Encoding.Unicode);   //create BinaryReader for the fileStream
            CFB = new CompoundFileBinary(fileReader);                                   //create CompoundFileBinary for fileReader
            if (CFB.CFHeaderIsOK) docIsOK = checkDOC();                                 //if file header is OK, we will check whether this file is a Word Binary File

            Fib.IsClear = true;                                                         //FIB is not read yet
            Clx.IsClear = true;                                                         //Clx is not read yet
        }

        /// <summary>
        /// Class constructor
        /// </summary>
        /// <param name="fileStream">FileStream of the opened file</param>
        protected internal docParser(FileStream fileStream)
        {
            if (fileStream == null) return;                                             //if no FileStream was specified we will just return

            BinaryReader fileReader = new BinaryReader(fileStream, Encoding.Unicode);   //create BinaryReader for the fileStream
            CFB = new CompoundFileBinary(fileReader);                                   //create CompoundFileBinary for fileReader
            if (CFB.CFHeaderIsOK) docIsOK = checkDOC();                                 //if file header is OK, we will check whether this file is a Word Binary File

            Fib.IsClear = true;                                                         //FIB is not read yet
            Clx.IsClear = true;                                                         //Clx is not read yet
        }
        #endregion

        #region Methods
        #region private
        /// <summary>
        /// Checks whether the file specified for this class is a Word Binary File
        /// </summary>
        /// <returns>true if file is a Word Binary File, false otherwise</returns>
        private bool checkDOC()
        {
            if (CFB == null) return false;                                                      //if CompoundFileBinary was not created we assume that file was not specified or file is wrong

            if (WDStream == null)                                                               //if WordDocument stream was not read from the CFB
            {
                string[] Paths = null;                                                          //paths to the found streams in the CFB
                uint[] StreamIds = null;                                                        //StreamIds of the found streams in the CFB
                if (!CFB.findStream("WordDocument", ref Paths, ref StreamIds)) return false;    //if no WordDocument stream found in CFB we assume this file is not Word Binary File

                uint streamID = 0;                                                              //WordDocument stream Id
                int i = 0;                                                                      //parameter of the following cicle
                for (i = 0; i < Paths.Length; i++)                                              //moving through all found streams
                {
                    if (Paths[i].IndexOf("ObjectPool") == -1)                                   //we're looking for the main WordDocument stream (not the ones that inserted to the file as OLE objects)
                    {
                        streamID = StreamIds[i];                                                //save StreamId of the main WordDocument stream
                        break;                                                                  //break the cicle
                    }
                }
                if (i == Paths.Length) return false;                                            //if we didn't find the main WordDocument stream in the CFB, assume that this file isn't Word Binary File
                WDStream = CFB.getStream(streamID);                                             //getting WordDocument stream from CFB
                WDStreamPath = Paths[i];                                                        //save path to WordDocument stream read from CFB
            }

            BinaryReader brWDStream = new BinaryReader(WDStream);                               //create BinaryReader for WDStream;
            WDStream.Seek(0, SeekOrigin.Begin);                                                 //seek to the beginning of the WDStream

            ushort wIdent = 0;                                                                  //{WD.FIB.FibBase.wIdent} Specifies that this is Word Binary File. (MUST: 0xA5EC) [off.: 0; len.: 2 bytes]
            wIdent = brWDStream.ReadUInt16();                                                   //read wIdent from WDStream

            if (wIdent == 0xA5EC) return true;                                                  //if wIdent equals 0xA5EC we assume this file is a Word Binary File

            return false;                                                                       //if we came here we assume the file is not a Word Binary File
        }

        /// <summary>
        /// Checks specified bit in specified number
        /// </summary>
        /// <param name="op">Number where to check bit</param>
        /// <param name="bit">Bit to be checked starting from 0</param>
        /// <returns>TRUE if bit is 1, FALSE if bit is 0</returns>
        private bool checkBit(uint op, int bit)
        {
            return ((op & (1 << bit)) == 0) ? false : true;
        }

        /// <summary>
        /// Checks specified bit in specified number
        /// </summary>
        /// <param name="op">Number where to check bit</param>
        /// <param name="bit">Bit to be checked starting from 0</param>
        /// <returns>TRUE if bit is 1, FALSE if bit is 0</returns>
        private bool checkBit(ushort op, int bit)
        {
            return ((op & (1 << bit)) == 0) ? false : true;
        }

        /// <summary>
        /// Checks specified bit in specified number
        /// </summary>
        /// <param name="op">Number where to check bit</param>
        /// <param name="bit">Bit to be checked starting from 0</param>
        /// <returns>TRUE if bit is 1, FALSE if bit is 0</returns>
        private bool checkBit(byte op, int bit)
        {
            return ((op & (1 << bit)) == 0) ? false : true;
        }

        /// <summary>
        /// Read FIB structure from WDStream
        /// </summary>
        /// <returns>TRUE if successfully read and found no errors in FIB</returns>
        private bool readFIB()
        {
            if (WDStream == null)                                                       //if WordDocument stream is not read from CFB
            {
                clearFIB();                                                             //clear all the FIB fields
                return false;                                                           //return false
            }

            BinaryReader brWDStream = new BinaryReader(WDStream);                       //create BinaryReader for WDStream

            //read Fib.base
            WDStream.Seek(10, SeekOrigin.Begin);                                        //seek WDStream to the offset of bitsAtoM
            Fib._base.AtoM = brWDStream.ReadBytes(2);                                   //read 2 bytes from WDStream
            Fib._base.fWhichTblStm = checkBit(Fib._base.AtoM[0], 6);                    //read fWhichTblStm

            //read Fib.fibRgLw
            WDStream.Seek(76, SeekOrigin.Begin);                                        //seek WDStream to the location of ccpText
            Fib.fibRgLw.ccpText = brWDStream.ReadInt32();                               //read ccpText from WDStream

            //read Fib.fibRgFcLcbBlob.fibRgFcLcb97
            WDStream.Seek(418, SeekOrigin.Begin);                                       //seek WDStream to the offset of fcClx
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.fcClx = brWDStream.ReadUInt32();            //read fcClx
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbClx = brWDStream.ReadUInt32();           //read lcbClx
            if (Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbClx <= 0) return false;              //lcbClx must be greater than zero. If it's not there are errors in FIB
            WDStream.Seek(250, SeekOrigin.Begin);                                       //seek WDStream to the offset of fcPlcfBteChpx
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.fcPlcfBteChpx = brWDStream.ReadUInt32();    //read fcPlcfBteChpx
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbPlcfBteChpx = brWDStream.ReadUInt32();   //read lcbPlcfBteChpx
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.fcPlcfBtePapx = brWDStream.ReadUInt32();    //read fcPlcBtePapx
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbPlcfBtePapx = brWDStream.ReadUInt32();   //read lcbPlcBtePapx

            Fib.IsClear = false;                                                        //FIB is not clear now

            return true;
        }

        /// <summary>
        /// Clear all fields of the memory representation of the FIB structure from WDStream
        /// </summary>
        private void clearFIB()
        {
            //just set all fields of Fib to default values            
            Fib._base.AtoM = null;
            Fib._base.fWhichTblStm = false;
            Fib.fibRgLw.ccpText = 0;
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.fcClx = 0;
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbClx = 0;
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.fcPlcfBteChpx = 0;
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbPlcfBteChpx = 0;
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.fcPlcfBtePapx = 0;
            Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbPlcfBtePapx = 0;

            //FIB is clear now
            Fib.IsClear = true;
        }

        /// <summary>
        /// Clear all fields of the memory representation of the Clx structure from TableStream 
        /// </summary>
        private void clearClx()
        {
            //just setting all the fields in Clx to its default values
            Clx.RgPrc = null;
            Clx.pcdt.clxt = 0;
            Clx.pcdt.lcb = 0;
            Clx.pcdt.plcPcd.aCP = null;
            Clx.pcdt.plcPcd.aPcd = null;

            //Clx is clear now
            Clx.IsClear = true;
        }

        /// <summary>
        /// Read Clx structure from TableStream
        /// </summary>
        /// <returns></returns>
        private bool readClx()
        {
            if (TableStream == null)                                    //if TableStream is not read from the file
            {
                clearClx();                                             //clear Clx
                return false;                                           //return false
            }

            if(Fib.IsClear)                                             //if FIB is not read the WordDocument stream
            {
                if (!readFIB())                                         //trying to read FIB, if couldn't
                {
                    clearClx();                                         //clear Clx
                    return false;                                       //return false
                }
            }

            BinaryReader brTableStream = new BinaryReader(TableStream); //create BinaryReader for TableStream

            TableStream.Seek(Fib.fibRgFcLcbBlob.fibRgFcLcb97.fcClx, SeekOrigin.Begin);          //seek TableStream to the offset of Clx
            byte[] clx = brTableStream.ReadBytes((int)Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbClx);  //read Clx from the Table stream

            //get RgPrc from clx
            MemoryStream msclx = new MemoryStream(clx);                                         //create MemoryStream from Clx
            BinaryReader brclx = new BinaryReader(msclx);                                       //create BinaryReader for msClx
            byte clxt = brclx.ReadByte();                                                       //read first byte from brClx. It'll be Clx.Prc.clxt or Clx.Pcdt.clxt
            while (clxt != 0x02)                                                                //while we haven't reached the beginning of the Pcdt
            {
                if (clxt != 0x01)                                                               //if that is not Prc then something is wrong with the Clx - we can't read it
                {
                    clearClx();                                                                 //clear Clx
                    return false;                                                               //and return false
                }
                if (Clx.RgPrc == null) Clx.RgPrc = new Prc[1];                                  //allocate memory for RgPrc
                else Array.Resize(ref Clx.RgPrc, Clx.RgPrc.Length + 1);                         //or reallocate it if already done
                int curRgPrcPos = Clx.RgPrc.Length - 1;                                         //current position in RgPrc array (index of the last item)
                Clx.RgPrc[curRgPrcPos].clxt = clxt;                                             //save clxt to RgPrc
                Clx.RgPrc[curRgPrcPos].data.cbGrpprl = brclx.ReadInt16();                       //read cdGrpprl from brClx
                if (Clx.RgPrc[curRgPrcPos].data.cbGrpprl > 0x3FA2)                              //if cbGrpprl is greater than 0x3FA2 then something is wrong with the Clx - we can't read it
                {
                    clearClx();                                                                 //clear Clx
                    return false;                                                               //return false
                }
                short leftToRead = Clx.RgPrc[curRgPrcPos].data.cbGrpprl;                        //number of bytes left to read from GrpPrl
                while (leftToRead > 0)                                                          //while there are bytes left unread in GrpPrl
                {
                    if (Clx.RgPrc[curRgPrcPos].data.GrpPrl == null)
                        Clx.RgPrc[curRgPrcPos].data.GrpPrl = new Prl[1];                        //allocate memory for GrpPrl
                    else
                        Array.Resize(ref Clx.RgPrc[curRgPrcPos].data.GrpPrl, 
                            Clx.RgPrc[curRgPrcPos].data.GrpPrl.Length + 1);                     //or reallocate it if already done
                    int curGrpPrlPos = Clx.RgPrc[curRgPrcPos].data.GrpPrl.Length - 1;           //current position in GrpPrl array (index of the last item)
                    short readBytes = readSprm(
                        ref Clx.RgPrc[curRgPrcPos].data.GrpPrl[curGrpPrlPos], 
                        ref brclx);                                                             //read current Prl from brClx
                    if (readBytes == 0)                                                         //if couldn't read Prl then Clx is corrupted and we can't read it
                    {
                        clearClx();                                                             //clear Clx
                        return false;                                                           //and return false
                    }
                    leftToRead -= readBytes;                                                    //decrease number of bytes left to read from GrpPrl by length of currently read Prl
                }
                clxt = brclx.ReadByte();                                                        //read the value of the next clxt
            }

            //read Pcdt and PlcPcd from Clx
            Clx.pcdt.clxt = clxt;                                                               //save last read clxt to Pcdt
            Clx.pcdt.lcb = brclx.ReadUInt32();                                                  //read Pcdt.lcb
            byte[] plcPcd = brclx.ReadBytes((int)Clx.pcdt.lcb);                                 //read PlcPcd from Clx
            brclx.Close();                                                                      //close brclx and msclx - we don't need them anymore

            //retrieve two arrays from PlcPcd: aCP and aPcd
            int n = ((int)Clx.pcdt.lcb - 4) / (8 + 4);                                          //number of data elements in PlcPcd (number of items in aPcd) (and number of items in aCP is (n+1))
            Clx.pcdt.plcPcd.aCP = new uint[n + 1];                                              //allocating memory for aCP - the array of CP elements that specifies the starting points of text ranges 
            Clx.pcdt.plcPcd.aPcd = new Pcd[n];                                                  //allocating memory for aPcd - the array of Pcds that specifies the location of text in WordDocument Stream
            MemoryStream msplcPcd = new MemoryStream(plcPcd);                                   //create MemoryStream for PlcPcd
            BinaryReader brplcPsd = new BinaryReader(msplcPcd);                                 //create BinaryReader for msPlcPcd
            for (int i = 0; i < (n + 1); i++)
                Clx.pcdt.plcPcd.aCP[i] = brplcPsd.ReadUInt32();                                 //read aCP from PlcPcd
            for (int i = 0; i < n; i++)                                                         //read aPcd from PlcPcd
            {
                msplcPcd.Seek(2, SeekOrigin.Current);                           //seek 2 bytes from current (to skip data that we do not need)
                Clx.pcdt.plcPcd.aPcd[i].fc.fc = brplcPsd.ReadUInt32();          //read FcCompressed.fc
                Clx.pcdt.plcPcd.aPcd[i].fc.fCompressed = 
                    checkBit(Clx.pcdt.plcPcd.aPcd[i].fc.fc, 30);                //retrieve FcCompressed.fCompressed
                Clx.pcdt.plcPcd.aPcd[i].fc.fc &= 0x3FFFFFFF;                    //use bitwise and to set bits 30 and 31 to 0 (because they are not for fc in FcCompressed)
                Clx.pcdt.plcPcd.aPcd[i].prm.Prm0_isprm = brplcPsd.ReadByte();   //read first byte of Prm structure, which is isprm
                Clx.pcdt.plcPcd.aPcd[i].prm.fComplex =
                    checkBit(Clx.pcdt.plcPcd.aPcd[i].prm.Prm0_isprm, 0);        //retrieve prm.fComplex bit
                if (Clx.pcdt.plcPcd.aPcd[i].prm.fComplex)                       //fComplex = 1 - we'll use Prm1
                {
                    msplcPcd.Seek(-1, SeekOrigin.Current);                      //seek -1 byte from current (to begin reading Prm1)
                    Clx.pcdt.plcPcd.aPcd[i].prm.Prm1_igrpprl = 
                        brplcPsd.ReadUInt16();                                  //read igrpprl
                    Clx.pcdt.plcPcd.aPcd[i].prm.Prm1_igrpprl >>= 1;             //shift igrpprl 1 bit to the left (that bit is fComplex - bit 0)
                    Clx.pcdt.plcPcd.aPcd[i].prm.Prm0_isprm = 0;                 //set Prm0.isprm to zero just for case
                    Clx.pcdt.plcPcd.aPcd[i].prm.Prm0_val = 0;                   //set Prm0.val to zero just for case
                }
                else                                                            //fComplex =0 - we'll use Prm0
                {
                    Clx.pcdt.plcPcd.aPcd[i].prm.Prm0_isprm >>= 1;               //shift isprm 1 bit to the left (that bit is fComplex - bit 0)
                    Clx.pcdt.plcPcd.aPcd[i].prm.Prm0_val = brplcPsd.ReadByte(); //read Prm0.val
                }
            }
            brplcPsd.Close();                                                   //we don't need brPlcPsd and msPlcPsd anymore and can close them

            Clx.IsClear = false;                                        //Clx is not clear now

            return true;                                                //return true
        }

        /// <summary>
        /// Read one Sprm from BinaryReader to Prl (stream of the BinaryReader must be set to the offset of Sprm)
        /// Moves the position in stream
        /// </summary>
        /// <param name="prl">Prl where Sprm is needed to be read</param>
        /// <param name="br">BinaryReader from which Sprm is needed to be read (set to the offset of Sprm</param>
        /// <returns>Number of bytes read (e.g. size of read Sprm in bytes) or zero if read is not successfull<returns>
        private short readSprm(ref Prl prl, ref BinaryReader br)
        {
            short bytesRead = 0;                                                    //number of bytes read from br

            if (br == null) return 0;                                               //if br is empty - we couldn't read Sprm

            //read sprm and interpret its fields
            ushort sprm = br.ReadUInt16();                                          //read Sprm as ushort
            bytesRead += 2;                                                         //increased number of read bytes
            prl.sprm.sprm = sprm;                                                   //save the 16-bits representation of currently read sprm
            prl.sprm.ispmd = (ushort)(sprm & 0x01FF);                               //use formula to calculate ispmd
            prl.sprm.fSpec = checkBit((ushort)(sprm / 512), 0);                     //use formula to calculate fSpec
            prl.sprm.sgc = (byte)((sprm / 1024) & 0x0007);                          //use formula to calculate sgc
            prl.sprm.spra = (byte)(sprm / 8192);                                    //use formula to calculet spra

            //read operand for the current sprm
            uint opSize = 1;                                                        //size of the operand
            switch (prl.sprm.spra)                                                  //switch between sizes of the current operand specified by spra
            {
                case 0: case 1: opSize = 1; break;                                  //spra = 0 OR 1: size is 1 byte
                case 2: case 4: case 5: opSize = 2; break;                          //spra = 2, 4 OR 5: size is 2 bytes
                case 7: opSize = 3; break;                                          //spra = 7: size is 3 bytes
                case 3: opSize = 4; break;                                          //spra = 3: size is 4 bytes
                case 6:                                                             //spra = 6: size depends on ispdm
                    if (prl.sprm.sgc == 5 && prl.sprm.ispmd == 0x08)                //if Sprm is sprmTDefTable
                    {
                        //operand is TDefTableOperand structure
                        opSize = br.ReadUInt16();                                   //{TDefTableOperand.cb} Number of bytes used by the remainder of this structure, incremented by 1 [2 bytes]
                        opSize++;                                                   //to get the full size of this operand
                        br.BaseStream.Seek(-2, SeekOrigin.Current);                 //seek stream back to the offset of the current operand
                    }
                    else if (prl.sprm.sgc == 1 && prl.sprm.ispmd == 0x15)           //if Sprm is sprmPChgTabs
                    {
                        //operand is PChgTabsOperand structure
                        opSize = br.ReadByte();                                     //{PChgTabsOperand.cb} Size in bytes of this operand (MUST >=2 AND <=255) [1 byte]
                        if (opSize < 2)                                             //if cb is less than 2, then we can't read this sprm
                        {
                            //set all fields of sprm to default values
                            prl.operand = null;
                            prl.sprm.sprm = 0;
                            prl.sprm.fSpec = false;
                            prl.sprm.spra = 0;
                            prl.sprm.sgc = 0;
                            prl.sprm.ispmd = 0;
                            //seek stream back to the starting offset
                            br.BaseStream.Seek(-(bytesRead + 1), SeekOrigin.Current);
                            //return 0
                            return 0;
                        }
                        if (opSize == 255)                                          //if cb == 255 
                        {
                            //reading the operand further
                            byte pctdcTabs = br.ReadByte();                         //{PChgTabsDelClose.cTabs} Number of records in rgdxaDel and rgdxaClose (MUST >=0 AND <=64) [1 byte]
                            br.BaseStream.Seek(pctdcTabs * 4, SeekOrigin.Current);  //seek stream to the offset of PChgTabsAdd
                            byte pctaTabs = br.ReadByte();                          //{PChgTabsAdd.cTabs} Number of records in rgdxaAdd and rgtbdAdd (MUST <=64) [1 byte]
                            opSize = (uint)(4 * pctdcTabs + 3 * pctdcTabs);         //calculated size of this operand without the first byte
                            br.BaseStream.Seek(-2, SeekOrigin.Current);             //seek stream back
                        }
                        opSize++;                                                   //to get the full size of this operand
                        br.BaseStream.Seek(-1, SeekOrigin.Current);                 //seek stream back to the offset of the current operand
                    }
                    else                                                            //for other Sprms
                    {
                        //read first byte of the operand
                        opSize = br.ReadByte();                                     //read size of this operand without the first byte
                        opSize++;                                                   //to get the full size of the operand
                        br.BaseStream.Seek(-1, SeekOrigin.Current);                 //seek stream back to the offset of this operand
                    }
                    break;
                default: break;                                                     //in other cases of spra assume size as 1 byte
            }

            prl.operand = br.ReadBytes((int)opSize);                                //{Prl.operand} Operand for the sprm [variable]
            bytesRead += (short)opSize;                                             //increase number of read bytes by the size of operand

            return bytesRead;                                                       //return number of bytes read
        }

        /// <summary>
        /// Read PlcBteChpx structure from Table stream
        /// </summary>
        /// <returns>TRUE if successfully read</returns>
        private bool readPlcBteChpx()
        {
            if(TableStream==null)                                                               //if TableStream is not read yet
            {
                //set all PlcBteChpx fields to default values and return false
                PlcBteChpx.aFC = null;
                PlcBteChpx.aPnBteChpx = null;
                return false;
            }

            BinaryReader brTableStream = new BinaryReader(TableStream);                         //create BinaryReader for TableStream

            TableStream.Seek(Fib.fibRgFcLcbBlob.fibRgFcLcb97.fcPlcfBteChpx, SeekOrigin.Begin);  //seek TableStream to the offset of PlcBteChpx
            byte[] plcBteChpx = brTableStream.ReadBytes(
                (int)Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbPlcfBteChpx);                           //read PlcBteChpx from the Table stream

            //retrieve two arrays from PlcBteChpx: aFC and aPnBteChpx
            int n = ((int)Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbPlcfBteChpx - 4) / (4 + 4);        //number of data elements in PlcBteChpx (number of items in aPnBteCHpx) (and number of items in aFC is (n+1))
            PlcBteChpx.aFC = new uint[n + 1];                                                   //allocate memory for aFC
            PlcBteChpx.aPnBteChpx = new PnFkpChpx[n];                                           //allocate memory for aPnBteChpx
            MemoryStream msplcBteChpx = new MemoryStream(plcBteChpx);                           //create MemoryStream for plcBteChpx
            BinaryReader brplcBteChpx = new BinaryReader(msplcBteChpx);                         //create BinaryReader for msplcBteChpx
            for (int i = 0; i < n + 1; i++) PlcBteChpx.aFC[i] = brplcBteChpx.ReadUInt32();      //read aFC from plcBteChpx
            for (int i = 0; i < n; i++)
            {
                PlcBteChpx.aPnBteChpx[i].pn = brplcBteChpx.ReadUInt32();                        //read aPnBteChpx.pn from PlcBteChpx
                PlcBteChpx.aPnBteChpx[i].pn &= 0x3FFFFF;                                        //use bitwise AND to drop 10 MSB in aPnBteChpx[i].pn - they're not used and must ne ignored
                PlcBteChpx.aPnBteChpx[i].unused = 0;                                            //initialise aPnBteChpx.unused to zero
            }
            brplcBteChpx.Close();                                                               //we don't need brPlcBteChpx & msPlcBteChpx anymore and can close them

            return true;                                                                        //return true
        }

        /// <summary>
        /// Read PlcBtePapx structure from Table stream
        /// </summary>
        /// <returns>TRUE if successfully read</returns>
        private bool readPlcBtePapx()
        {
            if (TableStream == null)                                                               //if TableStream is not read yet
            {
                //set all PlcBtePapx fields to default values and return false
                PlcBteChpx.aFC = null;
                PlcBteChpx.aPnBteChpx = null;
                return false;
            }

            BinaryReader brTableStream = new BinaryReader(TableStream);                         //create BinaryReader for TableStream

            TableStream.Seek(Fib.fibRgFcLcbBlob.fibRgFcLcb97.fcPlcfBtePapx, SeekOrigin.Begin);  //seek TableStream to the offset of PlcBtePapx
            byte[] plcBtePapx = brTableStream.ReadBytes(
                (int)Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbPlcfBtePapx);                           //read PlcBtePapx from the Table stream

            //retrieve two arrays from PlcBtePapx: aFC and aPnBtePapx
            int n = ((int)Fib.fibRgFcLcbBlob.fibRgFcLcb97.lcbPlcfBtePapx - 4) / (4 + 4);        //number of data elements in PlcBtePapx (number of items in aPnBtePapx) (and number of items in aFC is (n+1))
            PlcBtePapx.aFC = new uint[n + 1];                                                   //allocate memory for aFC
            PlcBtePapx.aPnBtePapx = new PnFkpChpx[n];                                           //allocate memory for aPnBtePapx
            MemoryStream msplcBtePapx = new MemoryStream(plcBtePapx);                           //create MemoryStream for plcBtePapx
            BinaryReader brplcBtePapx = new BinaryReader(msplcBtePapx);                         //create BinaryReader for msplcBtePapx
            for (int i = 0; i < n + 1; i++) PlcBtePapx.aFC[i] = brplcBtePapx.ReadUInt32();      //read aFC from plcBteChpx
            for (int i = 0; i < n; i++)
            {
                PlcBtePapx.aPnBtePapx[i].pn = brplcBtePapx.ReadUInt32();                        //read aPnBteChpx.pn from PlcBteChpx
                PlcBtePapx.aPnBtePapx[i].pn &= 0x3FFFFF;                                        //use bitwise AND to drop 10 MSB in aPnBtePapx[i].pn - they're not used and must ne ignored
                PlcBtePapx.aPnBtePapx[i].unused = 0;                                            //initialise aPnBtePapx.unused to zero
            }
            brplcBtePapx.Close();                                                               //we don't need brPlcBtePapx & msPlcBtePapx anymore and can close them

            return true;                                                                        //return true
        }

        /// <summary>
        /// Read whole array of ChpxFkp structures from WordDocument stream
        /// </summary>
        /// <returns>TRUE if successfully read</returns>
        private bool readChpxFkp()
        {
            if (WDStream == null)                                                           //if WDStream not read yet then we can't read CHpxFkp
            {
                aChpxFkp = null;                                                            //set aChpxFkp to null
                return false;                                                               //return false
            }

            aChpxFkp = new ChpxFkp[PlcBteChpx.aPnBteChpx.Length];                           //allocate memory for aChpxFkp

            BinaryReader brWDStream = new BinaryReader(WDStream);                           //create BinaryReader for WDStream

            for (int i = 0; i < aChpxFkp.Length; i++)                                       //moving through all the ChpxFkp-s
            {
                WDStream.Seek(PlcBteChpx.aPnBteChpx[i].pn * 512, SeekOrigin.Begin);         //seek WDStream to the beginning of current ChpxFkp
                byte[] aChpx = brWDStream.ReadBytes(512);                                   //read ChpxFkp

                MemoryStream msChpx = new MemoryStream(aChpx);                              //create MemoryStream for ChpxFkp
                BinaryReader brChpx = new BinaryReader(msChpx);                             //create BinaryReader for msChpx

                msChpx.Seek(-1, SeekOrigin.End);                                            //seek msChpx to the last byte (offset of crun)
                aChpxFkp[i].crun = brChpx.ReadByte();                                       //read crun
                aChpxFkp[i].rgfc = new uint[aChpxFkp[i].crun + 1];                          //allocate memory for aChpxFkp.rgfc
                aChpxFkp[i].rgb = new byte[aChpxFkp[i].crun];                               //allocate memory for aChpxFkp.rgb
                aChpxFkp[i].chpx = new Chpx[aChpxFkp[i].crun];                              //allocate memory for aChpxFkp.chpx

                msChpx.Seek(0, SeekOrigin.Begin);                                           //seek msChpx to the beginning
                for (int j = 0; j < aChpxFkp[i].crun + 1; j++)
                    aChpxFkp[i].rgfc[j] = brChpx.ReadUInt32();                              //read rgfc
                for (int j = 0; j < aChpxFkp[i].crun; j++)
                    aChpxFkp[i].rgb[j] = brChpx.ReadByte();                                 //read rgb

                for (int j = 0; j < aChpxFkp[i].crun; j++)                                  //moving through all the rgb-s
                {
                    if (aChpxFkp[i].rgb[j] == 0)                                            //if rgb == 0 then there is no Chpx associated with this element or rgb
                    {
                        aChpxFkp[i].chpx[j].cb = 0;                                         //then ChpxFkp.chpx.cb = 0
                        aChpxFkp[i].chpx[j].grpprl = null;                                  //and ChpxFkp.chpx.grpprl = null
                    }
                    else                                                                    //if Chpx.rgb != 0 then there is Chpx associated with this element of rgb
                    {
                        msChpx.Seek(aChpxFkp[i].rgb[j] * 2, SeekOrigin.Begin);              //seek msChpx to offset of the current chpx
                        aChpxFkp[i].chpx[j].cb = brChpx.ReadByte();                         //read Chpx.cb
                        if (aChpxFkp[i].chpx[j].cb != 0)                                    //if there is grpprl in this chpx then we'll read grpprl
                        {
                            int cbLeftBytes = aChpxFkp[i].chpx[j].cb;                     //number of bytes left to read from the current grpprl
                            while (cbLeftBytes > 0)                                         //reading while there are Prls unread
                            {
                                if (aChpxFkp[i].chpx[j].grpprl == null)
                                    aChpxFkp[i].chpx[j].grpprl = new Prl[1];                //allocate memory for grpprl
                                else
                                    Array.Resize(
                                        ref aChpxFkp[i].chpx[j].grpprl,
                                        aChpxFkp[i].chpx[j].grpprl.Length + 1);             //or reallocate it if already done
                                int curPrlPos = aChpxFkp[i].chpx[j].grpprl.Length - 1;      //current position in grpprl array (index of the last item)
                                short readBytes = readSprm(
                                        ref aChpxFkp[i].chpx[j].grpprl[curPrlPos],
                                        ref brChpx);                                        //read current Prl
                                if (readBytes == 0)                                         //if couldn't read current Prl then we cannot read aChpxFkp
                                {
                                    aChpxFkp = null;                                        //set aChpxFkp to null
                                    return false;                                           //and return false
                                }
                                cbLeftBytes -= readBytes;                                   //decrease number of bytes left to read
                            }
                        }
                        else aChpxFkp[i].chpx[j].grpprl = null;                             //if there is no grpprl in this chpx then chpx.grpprl = null
                    }
                }
                brChpx.Close();                                                             //close brChpx - we do not need it anymore
            }
            return true;                                                                    //return true
        }

        /// <summary>
        /// Read whole array of PapxFkp structures from WordDocument stream
        /// </summary>
        /// <returns>TRUE if successfully read</returns>
        private bool readPapxFkp()
        {
            if (WDStream == null)                                                                   //if WDStream not read yet then we can't read CHpxFkp
            {
                aPapxFkp = null;                                                                    //set aPapxFkp to null
                return false;                                                                       //return false
            }

            aPapxFkp = new PapxFkp[PlcBtePapx.aPnBtePapx.Length];                                   //allocate memory for aPapxFkp

            BinaryReader brWDStream = new BinaryReader(WDStream);                                   //create BinaryReader for WDStream

            for (int i = 0; i < aPapxFkp.Length; i++)                                               //moving through all the PapxFkp-s
            {
                WDStream.Seek(PlcBtePapx.aPnBtePapx[i].pn * 512, SeekOrigin.Begin);                 //seek WDStream to the beginning of current PapxFkp
                byte[] aPapx = brWDStream.ReadBytes(512);                                           //read PapxFkp

                MemoryStream msPapx = new MemoryStream(aPapx);                                      //create MemoryStream for PapxFkp
                BinaryReader brPapx = new BinaryReader(msPapx);                                     //create BinaryReader for msPapx

                msPapx.Seek(-1, SeekOrigin.End);                                                    //seek msPapx to the last byte (offset of cpara)
                aPapxFkp[i].cpara = brPapx.ReadByte();                                              //read cpara
                aPapxFkp[i].rgfc = new uint[aPapxFkp[i].cpara + 1];                                 //allocate memory for aPapxFkp.rgfc
                aPapxFkp[i].rgbx = new BxPap[aPapxFkp[i].cpara];                                    //allocate memory for aPapxFkp.rgbx
                aPapxFkp[i].papxInFkp = new PapxInFkp[aPapxFkp[i].cpara];                           //allocate memory for aPapxFkp.papxInFkp

                msPapx.Seek(0, SeekOrigin.Begin);                                                   //seek msPapx to the beginning
                for (int j = 0; j < aPapxFkp[i].cpara + 1; j++)
                    aPapxFkp[i].rgfc[j] = brPapx.ReadUInt32();                                      //read rgfc
                for (int j = 0; j < aPapxFkp[i].cpara; j++)
                {
                    aPapxFkp[i].rgbx[j].bOffset = brPapx.ReadByte();                                //read rgbx.bOffset
                    aPapxFkp[i].rgbx[j].reserved = brPapx.ReadBytes(12);                            //read rgbx.reserved
                }

                for (int j = 0; j < aPapxFkp[i].cpara; j++)                                         //moving through all the rgbx-s
                {
                    if (aPapxFkp[i].rgbx[j].bOffset == 0)                                           //if rgbx.bOffset == 0 then there is no PapxInFkp associated with this element or rgbx
                    {
                        aPapxFkp[i].papxInFkp[j].cb = 0;                                            //then PapxFkp.papxInFkp.cb = 0
                        aPapxFkp[i].papxInFkp[j].cb_ = 0;                                           //then PapxFkp.papxInFkp.cb_ = 0
                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.istd = 0;                             //then PapxFkp.papxInFkp.grpprlInPapx.istd = 0
                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.cbGrpprl = 0;                         //then PapxFkp.papxInFkp.grpprlInPapx.cbGrpprl = 0
                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl = null;                        //and ChpxFkp.chpx.grpprl = null
                    }
                    else                                                                            //if Papx.rgbx.bOffset != 0 then there is PaInFkp associated with this element of rgbx
                    {
                        msPapx.Seek(aPapxFkp[i].rgbx[j].bOffset * 2, SeekOrigin.Begin);             //seek msPapx to offset of the current papxInFkp
                        aPapxFkp[i].papxInFkp[j].cb = brPapx.ReadByte();                            //read papxInFkp.cb
                        int cb = aPapxFkp[i].papxInFkp[j].cb * 2 - 1;                               //calculate length of grpprlInPapx using cb (cb*2-1)
                        if (aPapxFkp[i].papxInFkp[j].cb == 0)                                       //if cb is zero
                        {
                            aPapxFkp[i].papxInFkp[j].cb_ = brPapx.ReadByte();                       //then read cb_
                            if (aPapxFkp[i].papxInFkp[j].cb_ < 1)                                   //if cb_ < 1 then we can't read aPapxFkp
                            {
                                aPapxFkp = null;                                                    //set aPapxFkp to null
                                return false;                                                       //ans return false
                            }
                            cb = aPapxFkp[i].papxInFkp[j].cb_ * 2;                                  //and calculate length of grpprlInPapx using cb_ (cb_*2)
                        }
                        if (cb != 0)                                                                //if there is grpprlInPapx in this papxInFkp then we'll read grpprlInPapx
                        {
                            aPapxFkp[i].papxInFkp[j].grpprlInPapx.istd = brPapx.ReadInt16();        //read istd
                            cb -= 2;                                                                //decrease number of bytes left to read
                            aPapxFkp[i].papxInFkp[j].grpprlInPapx.cbGrpprl = (short)cb;             //save cb value to cdGrpprl
                            while (cb > 0)                                                          //reading while there are GrpPrlAndIstd-s unread
                            {   
                                if (aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl == null)
                                    aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl = new Prl[1];      //allocate memory for grpprl
                                else
                                    Array.Resize(
                                        ref aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl,
                                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl.Length + 1);   //or reallocate it if already done
                                int curPrlPos =
                                    aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl.Length - 1;        //current position in grpprl array (index of the last item)
                                short readBytes = readSprm(
                                        ref aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl[curPrlPos],
                                        ref brPapx);                                                //read current Prl
                                if (aPapxFkp[i].papxInFkp[j].grpprlInPapx.
                                    grpprl[curPrlPos].sprm.sprm == 0x6646)                          //if current Sprm is sprmPHugePapx
                                {
                                    if (curPrlPos == 0)                                             //if this is the first Prl in array
                                    {
                                        //then we must read PrcData from DataStream to the GrpPrlAndIstd
                                        uint prcDataOffset =
                                            BitConverter.ToUInt32(
                                            aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl[0].operand, 
                                            0);                                                     //get offset of the PrcData in DataStream from the operand
                                        if (!readPrcDatafromDataStream(
                                            ref aPapxFkp[i].papxInFkp[j].grpprlInPapx,
                                            prcDataOffset))                                         //if couldn't retrieve PrcData from DataStream then we can't read aPapxFkp
                                        {
                                            aPapxFkp = null;                                        //set aPapxFkp to null
                                            return false;                                           //ans return false
                                        }
                                        break;                                                      //stop reading Prls from the current PapxFkp
                                    }
                                    else                                                            //if this Prl isn't first in this array
                                    {
                                        //then this Prl was added by mistake - we'll just clear it
                                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl[curPrlPos].operand = null;
                                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl[curPrlPos].sprm.fSpec = false;
                                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl[curPrlPos].sprm.ispmd = 0;
                                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl[curPrlPos].sprm.sgc = 0;
                                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl[curPrlPos].sprm.spra = 0;
                                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl[curPrlPos].sprm.sprm = 0;
                                    }
                                }
                                if (aPapxFkp[i].papxInFkp[j].grpprlInPapx.
                                    grpprl[curPrlPos].sprm.sprm == 0x646B)                          //if current Sprm is sprmPTableProps
                                {
                                    //then we must read PrcData from DataStream to the GrpPrlAndIstd
                                    uint prcDataOffset =
                                        BitConverter.ToUInt32(
                                        aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl[0].operand,
                                        0);                                                         //get offset of the PrcData in DataStream from the operand
                                    if (!readPrcDatafromDataStream(
                                        ref aPapxFkp[i].papxInFkp[j].grpprlInPapx,
                                        prcDataOffset))                                             //if couldn't retrieve PrcData from DataStream then we can't read aPapxFkp
                                    {
                                        aPapxFkp = null;                                            //set aPapxFkp to null
                                        return false;                                               //ans return false
                                    }
                                    break;                                                          //stop reading Prls from the current PapxFkp
                                }
                                if (readBytes == 0)                                                 //if couldn't read current Prl then we cannot read aPapxFkp
                                {
                                    aPapxFkp = null;                                                //set aPapxFkp to null
                                    return false;                                                   //and return false
                                }
                                cb -= readBytes;                                                    //decrease number of bytes left to read
                            }
                        }
                        else                                                                        //if there is no grpprl in this papxInFkp
                        {
                            aPapxFkp[i].papxInFkp[j].grpprlInPapx.grpprl = null;                    //then papxInFkp.grpprlInPapx.grpprl = null
                            aPapxFkp[i].papxInFkp[j].grpprlInPapx.cbGrpprl = 0;                     //and papxInFkp.grpprlInPapx.cbGrpprl = 0;
                        }
                    }
                }
                brPapx.Close();                                                                     //close brPapx - we do not need it anymore
            }
            return true;                                                                            //return true
        }

        /// <summary>
        /// Read PrcData structure from Data stream
        /// </summary>
        /// <param name="prcData">GrpPrlAndIstd structure to store PrcData in</param>
        /// <param name="offset">Offset of PrcData structure in the Data stream</param>
        /// <returns>TRUE if read successfully</returns>
        private bool readPrcDatafromDataStream(ref GrpPrlAndIstd prcData, uint offset)
        {
            if (DataStream == null)                                                             //if Data stream was not read from the CFB
            {
                //generate path to the Data stream and read it from CFB
                string Path = WDStreamPath.Substring(0, WDStreamPath.LastIndexOf('\\') + 1);    //Data stream should be located in the same storage as WordDocument stream
                Path += "Data";                                                                 //add the name of the Data stream to Path
                DataStream = CFB.getStream(Path);                                               //get Data stream from CFB
                if (DataStream == null) return false;                                           //if Data stream was not found we won't be able to read PrcData from it
            }

            BinaryReader brDataStream = new BinaryReader(DataStream);                           //create Binary Reader for DataStream

            DataStream.Seek(offset, SeekOrigin.Begin);                                          //seek DataStream to the offset of PrcData

            prcData.cbGrpprl = brDataStream.ReadInt16();                                        //read cdGrpprl
            if (prcData.cbGrpprl > 0x3FA2)                                                      //if cbGrpprl is greater than 0x3FA2 then we can't read the PrcData
            {
                //return false
                return false;
            }
            short leftToRead = prcData.cbGrpprl;                                                //number of bytes left to read from GrpPrl
            while (leftToRead > 0)                                                              //while there are bytes left unread in GrpPrl
            {
                if (prcData.grpprl == null)
                    prcData.grpprl = new Prl[1];                                                //allocate memory for GrpPrl
                else
                    Array.Resize(ref prcData.grpprl, prcData.grpprl.Length + 1);                //or reallocate it if already done
                int curGrpPrlPos = prcData.grpprl.Length - 1;                                   //current position in GrpPrl array (index of the last item)
                short readBytes = readSprm(
                    ref prcData.grpprl[curGrpPrlPos],
                    ref brDataStream);                                                          //read current Prl from brDataStream
                if (readBytes == 0)                                                             //if couldn't read Prl then we can't read PrcData
                {
                    //return false
                    return false;
                }
                leftToRead -= readBytes;                                                        //decrease number of bytes left to read from GrpPrl by length of currently read Prl
            }
            return true;                                                                        //return true
        }
        #endregion

        #region protected internal
        /// <summary>
        /// Retrieve text from the document
        /// </summary>
        /// <returns>String containing the document text (null if couldn't)</returns>
        protected internal string getText()
        {
            if (CFB == null) return null;                                                       //if CompoundFileBinary was not created there is nothing to read

            string[] Paths = null;                                                              //paths to the found streams in the CFB
            uint[] StreamIds = null;                                                            //StreamIds of the found streams in the CFB
            uint streamID = 0;                                                                  //Id of the actual stream to retrieved from the CFB
            string Path = null;                                                                 //Path of the actual stream to retrieved from the CFB

            if (WDStream == null)                                                               //if WordDocument stream was not read from the CFB
            {
                if (!CFB.findStream("WordDocument", ref Paths, ref StreamIds)) return null;     //if no WordDocument stream found in CFB there is nothing to read

                int i = 0;                                                                      //parameter of the following cicle
                for (i = 0; i < Paths.Length; i++)                                              //moving through all found streams
                {
                    if (Paths[i].IndexOf("ObjectPool") == -1)                                   //we're looking for the main WordDocument stream (not the ones that inserted to the file as OLE objects)
                    {
                        streamID = StreamIds[i];                                                //save StreamId of the main WordDocument stream
                        Path = Paths[i];                                                        //save Path of the main WordDocument
                        break;                                                                  //break the cicle
                    }
                }
                if (i == Paths.Length) return null;                                             //if we didn't find the main WordDocument stream in the CFB there is nothing to read
                WDStream = CFB.getStream(streamID);                                             //getting WordDocument stream from CFB
                WDStreamPath = Path;                                                            //save WordDocument stream path
            }

            BinaryReader brWDStream = new BinaryReader(WDStream);                               //create BinaryReader for WDStream

            if(Fib.IsClear)                                                                     //if FIB isn't read yet
                if (!readFIB()) return null;                                                    //try to read FIB from WDStream (return null if couldn't - we cannot retrieve text)

            if (TableStream == null)                                                            //if Table stream was not read from the CFB
            {
                //generate path to the Table stream and read it from CFB
                Path = WDStreamPath.Substring(0, WDStreamPath.LastIndexOf('\\') + 1);           //Table stream should be located in the same storage as WordDocument stream
                Path += (Fib._base.fWhichTblStm) ? "1Table" : "0Table";                         //add the name of the Table stream to Path depending on the value of the bit fWhichTblStm from FIB
                TableStream = CFB.getStream(Path);                                              //get Table stream from CFB
                if (TableStream == null) return null;                                           //if Table stream was not found we won't be able to read text from file
            }

            BinaryReader brTableStream = new BinaryReader(TableStream);                         //create BinaryReader for TableStream

            if(Clx.IsClear)                                                                     //is Clx isn't read yet
                if (!readClx()) return null;                                                    //try to read Clx from TableStream (return null if couldn't - we cannot retrieve text)

            if (!readPlcBteChpx()) return null;                                                 //try to read PlcBteChpx from the TableStream (return null if couldn't - we cannot retrieve text)

            if (!readPlcBtePapx()) return null;                                                 //try to read PlcBtePapx from the TableStream (return null if couldn't - we cannot retrieve text)

            if (!readChpxFkp()) return null;                                                    //try to read aChpxFkp from the WDStream (return null if couldn't - we cannot retrieve text)

            if (!readPapxFkp()) return null;                                                    //try to read aPapxFkp from the WDStream (return null if couldn't - we cannot retrieve text)

            //reading text from WordDocument stream
            //
            //NOTE: Text in WordDocument stream is splitted on blocks.
            //      Encoding used in each block is described by FcCompressed.fCompressed bit.
            //
            //      If it is false (zero) in aPcd[N] then Unicode is used for text-block number N, each character occupies 2 bytes, 
            //      the text-block is located at offset FcCompressed.fc and number of characters in this block is aCP[N+1]-aCP[N].
            //
            //      If it is true (one) in aPcd[N] then ANSI is used for text-block number N, each character occupies 1 byte,
            //      the text-block is located at offset (FcCompressed.fc/2) and number of characters in this block is aCP[N+1]-aCP[N].
            //
            //      There is one more nuance about ANSI text-blocks: there is a list of mapped byte values, that are used not as ANSI
            //      characters but as wildcards for some Unicode characters. I use a static class MappedToUnicode for them in which
            //      I declared a Dictionary collection. The Key in every pair of that Dictionary is ANSI 1 byte value and the Value
            //      is a Unicode character.
            //
            string docText = "";                                                                //buffer for the text retrieved from the document
            int textLen = Fib.fibRgLw.ccpText;                                                  //length of the text in MainDocument
            int charCount = Fib.fibRgLw.ccpText;                                                //counter of characters read from the MainDocument
            for (int i = 0; i < Clx.pcdt.plcPcd.aPcd?.Length; i++)                              //moving through all Pcds in aPcd
            {
                string readText = "";                                                           //current text-block read from the WordDocument stream
                byte[] readBytes = null;                                                        //current bytes-block read from the WordDocument stream
                uint fc = 0;                                                                    //offset of the current character in WordDocument stream
                uint dfc = 0;                                                                   //size in bytes of the current character
                if (Clx.pcdt.plcPcd.aPcd[i].fc.fCompressed)                                     //if fCompressed is true we will read ANSI
                {
                    fc = Clx.pcdt.plcPcd.aPcd[i].fc.fc / 2;                                     //offset of the current text-block in WordDocument stream
                    dfc = 1;                                                                    //size of one character for current text-block is 1
                    WDStream.Seek(fc, SeekOrigin.Begin);                                        //seek to needed offset in WordDocument stream
                    readBytes = brWDStream.ReadBytes(
                        (int)(Clx.pcdt.plcPcd.aCP[i + 1] - 
                        Clx.pcdt.plcPcd.aCP[i]));                                               //read current bytes-block from WordDocument stream
                    readText = Encoding.Default.GetString(readBytes);                           //convert ANSI bytes to Unicode string
                    for (int j = 0; j < readBytes.Length; j++)                                  //moving through all the read bytes
                    {
                        char tmpChar;                                                           //temporary character
                        if (MappedToUnicode.values.TryGetValue(readBytes[j], out tmpChar))      //trying to find current ANSI byte amidst the MappedToUnicode values
                            readText = readText.Substring(0, j) + 
                                tmpChar + 
                                readText.Substring(j + 1);                                      //if found, replace corresponding character in Unicode string with the one from MappedToUnicode
                    }
                }
                else                                                                            //if fCompressed is false we will read Unicode
                {
                    fc = Clx.pcdt.plcPcd.aPcd[i].fc.fc;                                         //offset of the current text-block in WordDocument stream
                    dfc = 2;                                                                    //size of one character for current text-block is 2
                    WDStream.Seek(fc, SeekOrigin.Begin);                                        //seek to needed offset in WordDocument stream
                    readBytes = brWDStream.ReadBytes(
                        (int)(2 * (Clx.pcdt.plcPcd.aCP[i + 1] - 
                        Clx.pcdt.plcPcd.aCP[i])));                                              //read current bytes-block from WordDocument stream
                    readText = Encoding.Unicode.GetString(readBytes);                           //converted read bytes-block to Unicode string
                }

                //compose the result text string using properties of the read characters
                foreach (char ch in readText)                                                   //moving through all characters in the current text-block
                {
                    bool isCharVisible = true;                                                  //is current character visible in the text or not

                    //determine visibility of current character by paragraph properties (PlcBtePapx & aPapxFkp)
                    int nFC = 0;                                                                //index in aFC corresponding to the current character offset
                    int nrgfc = 0;                                                              //index in PapxFkp.rgfc corresponding to the current character offset
                    for (nFC = PlcBtePapx.aFC.Length - 1; nFC >= 0; nFC--)                      //moving through all the items in aFC
                        if (PlcBtePapx.aFC[nFC] <= fc) break;                                   //looking for the index in aFC
                    if (nFC < PlcBtePapx.aFC.Length - 1)                                        //if index is found (fc is a valid offset)
                    {
                        for (nrgfc = aPapxFkp[nFC].rgfc.Length - 1; nrgfc >= 0; nrgfc--)        //moving throuth all the items in the current PapxFkp.rgfc
                            if (aPapxFkp[nFC].rgfc[nrgfc] <= fc) break;                         //lokking for the index in PapxFkp.rgfc
                        if (nrgfc < aPapxFkp[nFC].rgfc.Length - 1)                              //if index is found (fc is a valid offset)
                        {
                            isCharVisible = PropSPRM.IsVisible(
                                aPapxFkp[nFC].papxInFkp[nrgfc].grpprlInPapx.grpprl,
                                ch);                                                            //check visibility of the current character
                        }
                    }

                    //determine visibility of current character by direct characters properties (PlcBteChpx & aChpxFkp)
                    nFC = 0;                                                                    //index in aFC corresponding to the current character offset
                    nrgfc = 0;                                                                  //index in ChpxFkp.rgfc corresponding to the current character offset
                    for (nFC = PlcBteChpx.aFC.Length - 1; nFC >= 0; nFC--)                      //moving through all the items in aFC
                        if (PlcBteChpx.aFC[nFC] <= fc) break;                                   //looking for the index in aFC
                    if (nFC < PlcBteChpx.aFC.Length - 1)                                        //if index is found (fc is a valid offset)
                    {
                        for (nrgfc = aChpxFkp[nFC].rgfc.Length - 1; nrgfc >= 0; nrgfc--)        //moving throuth all the items in the current ChpxFkp.rgfc
                            if (aChpxFkp[nFC].rgfc[nrgfc] <= fc) break;                         //lokking for the index in ChpxFkp.rgfc
                        if (nrgfc < aChpxFkp[nFC].rgfc.Length - 1)                              //if index is found (fc is a valid offset)
                        {
                            isCharVisible = PropSPRM.IsVisible(
                                aChpxFkp[nFC].chpx[nrgfc].grpprl,
                                ch);                                                            //check visibility of the current character
                        }
                    }

                    //use information of visibility either to append character to the result string or not
                    if (isCharVisible) docText += ch;                                           //if current char is visible, add it to the result text string
                    else textLen--;                                                             //if current char isn't visible we'll drop it and decrease the number of characters in MainDocument by 1
                    if ((--charCount) <= 0) break;                                              //decrease counter of characters read and break if we've read all characters from the MainDocument
                    fc += dfc;                                                                  //go to the next character and next offset
                }
                if (charCount <= 0) break;                                                      //break if we've read all characters from the MainDocument
            }

            docText = docText.Substring(0, textLen);                                            //cut buffer to the length of the text in MainDocument

            return docText;                                                                     //return the retrieved text
        }

        /// <summary>
        /// Close DOC-file
        /// </summary>
        protected internal void closeDOC()
        {
            //close cfb readers
            CFB.closeReader();

            //clear everything
            clearFIB();
            clearClx();
            CFB = null;
            WDStream = null;
            WDStreamPath = null;
            TableStream = null;
            PlcBteChpx.aFC = null;
            PlcBteChpx.aPnBteChpx = null;
            PlcBtePapx.aFC = null;
            PlcBtePapx.aPnBtePapx = null;
            aChpxFkp = null;
            aPapxFkp = null;
            docIsOK = false;
        }
        #endregion
        #endregion
    }
}
