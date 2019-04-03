using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices.ComTypes;

namespace WordReader
{
    class docParser
    {
        #region Подклассы
        /// <summary>
        /// OLE Compound File Binary class
        /// </summary>
        private class CompoundFileBinary
        {
            #region Структуры
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

            #region Поля
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

            #region Конструкторы
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

            #region Методы
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
                string newPath = curPath + "\\" + curName;
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
            #endregion
            #endregion
        }
        #endregion

        #region Поля
        #region private
        private CompoundFileBinary CFB = null;      //class for reading the Compound Binary File
        private MemoryStream WDStream = null;       //WordDocument stream
        #endregion

        #region protected internal
        /// <summary>
        /// True is file exists, is OK and is a Word Binary File
        /// </summary>
        protected internal bool docIsOK = false;
        #endregion
        #endregion

        #region Свойства

        #endregion

        #region Конструкторы
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
        }
        #endregion

        #region Методы
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
            }

            BinaryReader brWDStream = new BinaryReader(WDStream);                               //create BinaryReader for WDStream;
            WDStream.Seek(0, SeekOrigin.Begin);                                                 //seek to the beginning of the WDStream

            ushort wIdent = 0;                                                                  //{FIB.FibBase.wIdent} Specifies that this is Word Binary File. (MUST: 0xA5EC) [off.: 0; len.: 2 bytes]
            wIdent = brWDStream.ReadUInt16();                                                   //read wIdent from WDStream

            if (wIdent == 0xA5EC) return true;                                                  //if wIdent equals 0xA5EC we assume this file is a Word Binary File

            return false;                                                                       //if we came here we assume the file is not a Word Binary File
        }
        #endregion

        #region protected internal

        #endregion
        #endregion
    }
}
