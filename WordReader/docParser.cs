using System;
using System.IO;
using System.Collections;
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
        private class CompoundFileBinary  //OLE Compound File Binary
        {
            #region Структуры
            private struct SpecialValues    //Reserved special values
            {
                internal const uint DIFSECT = 0xFFFFFFFC;       //Specifies a DIFAT sector in the FAT
                internal const uint FATSECT = 0xFFFFFFFD;       //Specifies a FAT sector in the FAT
                internal const uint ENDOFCHAIN = 0xFFFFFFFE;    //End of a linked chain of sectors
                internal const uint FREESECT = 0xFFFFFFFF;      //Specifies an unallocated sector in the FAT, Mini FAT or DIFAT
                internal const uint NOSTREAM = 0xFFFFFFFF;      //Terminator or empty pointer if Directory Entry
            }

            private struct CompoundFileHeader //Compound File Header structure
            {
                internal byte[] Signature;           //Header Signature (MUST:  0xD0CF11E0A1B11AE1) [8 bytes]
                internal byte[] CLSID;               //Header CLSID (MUST: all zeros) [16 bytes]
                internal byte[] MinorVersion;        //Minor Version (SHOULD: 0x3E00, if MajorVersion is 0x0300 or 0x0400) [2 bytes]
                internal byte[] MajorVersion;        //Major Version (MUST: 0x0300 (version 3) or 0x0400 (version 4)) [2 bytes]
                internal byte[] ByteOrder;           //Byte order (MUST: 0xFEFF) - little-endian [2 bytes]
                internal byte[] SectorShift;         //Sector shift (MUST: 0x0009 (if major version is 3) or 0x000C (if major version is 4)) [2 bytes]
                internal byte[] MiniSectorShift;     //Mini sector shift (sector size of the Mini Stream) (MUST: 0x0006) [2 bytes]
                internal byte[] Reserved;            //Reserved [6 bytes]
                internal uint NumDirSectors;         //Number of Directory sectors (MUST: 0x0 if major version is 3) [1 uint = 4 bytes]
                internal uint NumFATSectors;         //Number of FAT sectors [1 uint = 4 bytes]
                internal uint FirstDirSectorLoc;     //First directory sector location - starting sector nmber for directory stream [1 uint = 4 bytes]
                internal uint TransSignNum;          //Transaction signature number - how many times files was saved by implementation [1 uint = 4 bytes]
                internal uint MiniStreamCutoffSize;  //Max size of user-defined data stream (MUST: 0x00001000 = 4096) [1 uint = 4 bytes]
                internal uint FirstMiniFATSectorLoc; //First mini FAT sector location - starting sector number for mini FAT [1 uint = 4 bytes]
                internal uint NumMiniFATSectors;     //Number of mini FAT sectors [1 uint = 4 bytes]
                internal uint FirstDIFATSectorLoc;   //First DIFAT sector location - starting sector number for DIFAT [1 uint = 4 bytes]
                internal uint NumDIFATSectors;       //Number of DIFAT sectors [1 uint = 4 bytes]
                internal uint[] DIFAT;               //The first 109 FAT sector locations [109 uint = 436 bytes]
                //NOTE: for major version 3 CFHeader size is 512 bytes.
                //NOTE: for major version 4 CFHeader size is 4096 bytes, so the remaining part (3584 bytes) if filled with zeros
            }

            private struct DirectoryEntry   //Compound File Directory Entry structure
            {
                internal string Name;           //Directory Entry Name [64 bytes]
                internal uint NameLength;       //Directory Entry Name in bytes (MUST: <=64) [2 bytes]
                internal byte ObjectType;       //Object Type of the current directory entry
                                                //(MUST: 0x00 (unknown or unallocated, 0x01 (Storage object), 0x02 (Stream object) OR 0x05 (Root Storage object)
                                                //[1 byte]
                internal byte ColorFlag;        //Color flag of the current directory entry (MUST: 0x00 (red), 0x01 (black)) [1 byte]
                internal uint LeftSibling;      //Left Sibling stream ID (MUST: 0xFFFFFFFF if there is no left sibling) [4 bytes]
                internal uint RightSibling;     //Right Sibling stream ID (MUST: 0xFFFFFFFF if there is no right sibling) [4 bytes]
                internal uint Child;            //Child object stream ID (MUST: 0xFFFFFFFF if there is no child objects) [4 bytes]
                internal byte[] CLSID;          //Object class GUID, if current entry is for a storage object or root storage object
                                                //(MUST: all zeros for a stream object. MAY: all zeros for storage object or root storage object,
                                                //thus indicating that no object class is associated with the storage)
                                                //[16 bytes]
                internal byte[] StateBits;      //User-defined flags if current entry is for a storage object or root storage object
                                                //(SHOULD: all zeros for a stream object)
                                                //[4 bytes]
                internal long CreationTime; //Creation Time for a storage object (MUST: all zeros for a stream object OR root storage object) [8 bytes]
                internal long ModifiedTime; //Modification Time for a storage object (MUST: all zeros for a stream object. MAY: all zeros for a root storage object) [8 bytes]
                internal uint StartSectorLoc;   //Starting Sector Location  if this is a stream object (MUST: all zeros for a storage object.
                                                //MUST: first sector of the mini stream for a root storage object if the mini stream exists)
                                                //[4 bytes]
                internal ulong StreamSizeV4;    //Size of the user-defined data if this is a stream object. Size of the mini stream for a root storage object
                                                //(MUST: all zeros for a storage object)
                                                //[8 bytes]                                                
                internal uint StreamSizeV3;     //NOTE: THIS FIELD IS NOT IN REAL COMPOUND FILE DIRECTORY ENTRY STRUCTURE! I ADDED IT JUST FOR MY OWN CONVENIENCE!
                                                //Same as StreamSizeV4, but used for version 3 compound files. That is StreamSizeV4 without most significant 32 bits.
            }

            private struct FolderTreeEntry //структура записи для стека отображения дерева папок
            {
                internal int TreeLevel; //уровень в дереве (у RootEntry равен 0, далее каждый проход по Child добавляет 1)
                internal string Name;   //имя записи
                internal string Parent; //имя родителя
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

            #endregion
            #endregion

            #region Свойства

            #endregion

            #region Конструкторы
            protected internal CompoundFileBinary(BinaryReader reader)
            {
                fileReader = reader;
                readCFHeader();
                readDIFAT();
                readFAT();
                readminiFAT();
                readDEArray();
            }
            #endregion

            #region Методы
            #region private            
            private void showFCHeader() //вывод CFHeader на консоль
            {
                //перепишем все байты заголовка в один массив
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

                //выводим полученный массив
                showBytesInHEX(Output, "Compound file header", "end of header");
            }

            private void showBytesInHEX(byte[] Output, string title = "", string ending = "") //вывод байтового массива в виде HEX
                                                                                              //title - заголовок перед выводом, ending - строка после вывода
            {
                int byteNumber = 0; //кол-во выведенных байт
                Console.WriteLine("\t" + title);    //заголовок
                Console.Write($"{byteNumber:X6}: ");    //номер первого байта в текущей строке
                foreach (byte o in Output)  //побежали по всем байтам
                {
                    if (byteNumber != 0 && (byteNumber % 16) == 0)  //вывели 16 байт
                    {
                        Console.WriteLine();    //начнем новую строку
                        Console.Write($"{byteNumber:X6}: ");    //в следующей строке выведем номер первого байта
                    }
                    Console.Write($"{o:X2}");   //выведем текущий байт
                    byteNumber++;       //увеличим счетчик выведенных байт
                    if (byteNumber % 2 == 0) Console.Write(" ");    //через каждые два байта выводим пробел
                }

                Console.WriteLine("\n\t" + ending); //строка после вывода
            }

            private bool readCFHeader() //читает Compound file header из fileReader и проверяет его на адекватность
                                        //вернет true, если Compound file header без ошибок
            {
                fileReader.BaseStream.Seek(0, SeekOrigin.Begin);    //перемотка на начало файла

                //читаем заголовок
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
                                
                //эталонные (MUST) значения полей заголовка
                byte[] signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
                byte[] minorVersion = { 0x3E, 0x00 };
                byte[][] majorVersion = { new byte[] { 0x03, 0x00 }, new byte[] { 0x04, 0x00 } };
                byte[] byteOrder = { 0xFE, 0xFF };
                byte[][] sectorShift = { new byte[] { 0x09, 0x00 }, new byte[] { 0x0C, 0x00 } };
                byte[] miniSectorShift = { 0x06, 0x00 };

                //проверяем на ошибки
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

                return true;    //ошибок не обнаружено
            }

            private void readDIFAT()    //читаем полный массив DIFAT из fileReader
            {
                //копируем DIFAT из CFHeader
                for (int i = 0; i < CFHeader.DIFAT.Length; i++)
                {
                    if (CFHeader.DIFAT[i] != SpecialValues.FREESECT)
                    {
                        //выделение памяти
                        if (DIFAT == null) DIFAT = new uint[1];
                        else Array.Resize(ref DIFAT, DIFAT.Length + 1);
                        //копирование значения
                        DIFAT[i] = CFHeader.DIFAT[i];
                    }
                    else break;
                }

                //если DIFAT секторов в файле нет, то больше копировать нечего
                if (CFHeader.NumDIFATSectors == 0 || CFHeader.FirstDIFATSectorLoc == SpecialValues.ENDOFCHAIN) return;

                //ищем и копируем данные из DIFAT секторов
                uint numOfDIFATSectors = CFHeader.NumDIFATSectors;  //кол-во DIFAT секторов в файле
                uint curDIFATSEctorLoc = CFHeader.FirstDIFATSectorLoc; //адрес текущего сектора
                uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));    //размер сектора в файле
                int numEntriesInDIFAT = (int)(sectorSize - 4) / 4;  //кол-во записей DIFAT в одном DIFAT секторе

                //пока еще есть непрочитанные DIFAT сектора и пока не дошли до конца цепочки DIFAT секторов
                while (numOfDIFATSectors > 0 && curDIFATSEctorLoc != SpecialValues.ENDOFCHAIN)
                {
                    uint sectorOffset = (curDIFATSEctorLoc + 1) * sectorSize;  //номер первого байта текущего DIFAT сектора в файле
                    fileReader.BaseStream.Seek(sectorOffset, SeekOrigin.Begin); //перемотали файл на нужную позицию
                    for (int i = 0; i < numEntriesInDIFAT; i++) //читаем из файла все записи текущего сектора DIFAT, кроме последней
                    {
                        uint tmp = fileReader.ReadUInt32(); //читаем одну запись DIFAT из файла
                        if (tmp != SpecialValues.FREESECT) //если она не пустая
                        {
                            //перевыделим память
                            Array.Resize(ref DIFAT, DIFAT.Length + 1);
                            //сохраним, что прочитали
                            DIFAT[DIFAT.Length - 1] = tmp;
                        }
                    }
                    curDIFATSEctorLoc = fileReader.ReadUInt32();    //читаем номер следующего DIFAT сектора
                    numOfDIFATSectors--;    //закончили чтение текущего сектора
                }
            }

            private void readFAT()  //чтение FAT из fileReader
            {
                uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));    //размер сектора в файле
                int numEntriesInFAT = (int)(sectorSize) / 4;  //кол-во записей FAT в одном FAT секторе

                //выделение памяти
                FAT = new uint[CFHeader.NumFATSectors * numEntriesInFAT];

                for (int i = 0; i < DIFAT.Length; i++)
                {
                    uint sectorOffset = (DIFAT[i] + 1) * sectorSize;    //номер первого байта текущего FAT сектора в файле
                    fileReader.BaseStream.Seek(sectorOffset, SeekOrigin.Begin); //перемотали файл на нужную позицию
                    //читаем данные
                    for (int j = i * numEntriesInFAT; j < (i + 1) * numEntriesInFAT; j++)
                        FAT[j] = fileReader.ReadUInt32();
                }
            }

            private void readminiFAT()  //чтение полной таблицы miniFAT из CFHeader и FAT
            {
                if (CFHeader.NumMiniFATSectors == 0)    //если в файле нет miniFat секторов
                {
                    miniFAT = null;
                    return;
                }

                uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));    //размер сектора в файле
                uint miniFATSectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.MiniSectorShift, 0));    //размер сектора в miniStream
                uint mfEntriesPerSector = sectorSize / 4;   //кол-во записей miniFAT в одном секторе файла
                uint numMiniFATEntries = CFHeader.NumMiniFATSectors * mfEntriesPerSector;    //кол-во всех записей в miniFAT
                miniFAT = new uint[numMiniFATEntries]; //выделили память под miniFAT
                uint currentminiFATsector = CFHeader.FirstMiniFATSectorLoc;  //номер текущего сектора miniFAT
                int posInMiniFAT = 0;   //текущее положение в массиве miniFAT

                while (currentminiFATsector != SpecialValues.ENDOFCHAIN)    //пока не достигли конца цепочки FAT секторов
                {
                    uint fileOffset = (currentminiFATsector + 1) * sectorSize;  //положение текущего сектора в файле
                    fileReader.BaseStream.Seek(fileOffset, SeekOrigin.Begin);   //перемотали файл
                    byte[] readSector = fileReader.ReadBytes((int)sectorSize);  //прочитали текущий сектор
                    MemoryStream ms = new MemoryStream(readSector); //сделали из прочитанного сектора MemoryStream
                    BinaryReader br = new BinaryReader(ms);     //создали читалку для прочитанного сектора
                    for (int i = 0; i < mfEntriesPerSector; i++) miniFAT[posInMiniFAT + i] = br.ReadUInt32();   //читаем записи miniFAT из MemoryStream
                    posInMiniFAT += (int)mfEntriesPerSector;    //увеличили текущее положение в массиве miniFAT
                    currentminiFATsector = FAT[currentminiFATsector];   //перешли к следующему сектору
                }
            }

            private void readDEArray()  //чтение Directory Entry Array из fileReader
            {
                uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));    //размер сектора в файле
                int numDirEntries = (int)(sectorSize) / 128;    //кол-во directory entry в одном секторе (4 для version 3 или 32 для version 4) - размер одной entry = 128 bytes
                uint currentDirSector = CFHeader.FirstDirSectorLoc; //текущий сектор с Directory Entry
                int curDirSectorOrder = 0;  //номер текущего сектора по порядку
                while (currentDirSector != SpecialValues.ENDOFCHAIN)    //пока не разберем всю цепочку directory stream
                {
                    //выделение памяти
                    if (DEArray == null) DEArray = new DirectoryEntry[numDirEntries];   //первый сектор с Directory entry
                    else Array.Resize(ref DEArray, DEArray.Length + numDirEntries);     //все последующие сектора
                    uint sectorOffset = (currentDirSector + 1) * sectorSize;            //номер первого байта текущего сектора в файле
                    fileReader.BaseStream.Seek(sectorOffset, SeekOrigin.Begin);     //перемотали файл

                    for (int i = curDirSectorOrder * numDirEntries; i < (curDirSectorOrder + 1) * numDirEntries; i++)   //пробегаем все directory entry в текущем секторе
                    {
                        //читаем данные текущей directory entry
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

                        //вычисляем StreamSizeV3
                        DEArray[i].StreamSizeV3 = Convert.ToUInt32(DEArray[i].StreamSizeV4);
                    }

                    //находим номер следующего сектора с directory entry
                    currentDirSector = FAT[currentDirSector];
                    curDirSectorOrder++;    //увеличили номер текущего сектора по порядку
                }
            }

            private void buildFolderTree(uint Id, ref FolderTreeEntry[] FTE) //строим дерево папок для вывода на экран
                                                                             //Id - номер текущей записи Directory Entry
                                                                             //FTE - массив с деревом папок
            {
                //возврат, если попали в NOSTREAM
                if (Id == SpecialValues.NOSTREAM) return;

                //заполнение имени текущей записи
                int curFTE = FTE.Length - 1;
                FTE[curFTE].Name = DEArray[Id].Name.Substring(0, DEArray[Id].Name.IndexOf('\0'));
                FTE[curFTE].Name += (DEArray[Id].ObjectType == 0x00) ? " <unknown>" : "";
                FTE[curFTE].Name += (DEArray[Id].ObjectType == 0x01) ? " <storage>" : "";
                FTE[curFTE].Name += (DEArray[Id].ObjectType == 0x02) ? " <stream>" : "";

                //идем по Child, если он есть
                if (DEArray[Id].Child != SpecialValues.NOSTREAM)
                {
                    //перевыделим память и заполним данные по Child
                    Array.Resize(ref FTE, FTE.Length + 1);
                    FTE[FTE.Length - 1].TreeLevel = FTE[curFTE].TreeLevel + 1;
                    FTE[FTE.Length - 1].Parent = FTE[curFTE].Name;
                    buildFolderTree(DEArray[Id].Child, ref FTE);
                }

                //идем по Left, если он есть
                if (DEArray[Id].LeftSibling != SpecialValues.NOSTREAM)
                {
                    //перевыделим память и заполним данные по Left
                    Array.Resize(ref FTE, FTE.Length + 1);
                    FTE[FTE.Length - 1].TreeLevel = FTE[curFTE].TreeLevel;
                    FTE[FTE.Length - 1].Parent = FTE[curFTE].Parent;
                    buildFolderTree(DEArray[Id].LeftSibling, ref FTE);
                }

                //идем по Right, если он есть
                if (DEArray[Id].RightSibling != SpecialValues.NOSTREAM)
                {
                    //перевыделим память и заполним данные по Right
                    Array.Resize(ref FTE, FTE.Length + 1);
                    FTE[FTE.Length - 1].TreeLevel = FTE[curFTE].TreeLevel;
                    FTE[FTE.Length - 1].Parent = FTE[curFTE].Parent;
                    buildFolderTree(DEArray[Id].RightSibling, ref FTE);
                }
            }

            private void findInDEArray(uint Id, string curPath, string Name, ref string[] Paths, ref uint[] StreamIds)  //поиск потока в файле по заданному имени Name
                                                                                                                        //в Paths положит массив путей до найденных потоков (или null, если ничего не нашел)
                                                                                                                        //в StreamIds положит StreamId найденных потоков, или null, если ничего не нашел
                                                                                                                        //Id - Id текущей записи в DEArray
                                                                                                                        //curPath - на данный момент собираемый путь
            {
                if (Id == SpecialValues.NOSTREAM) return;   //если попали в NOSTREAM

                //проверяем текущую Directory Entry
                string curName = DEArray[Id].Name.Substring(0, DEArray[Id].Name.IndexOf('\0'));
                if (curName.CompareTo(Name) == 0)  //если имена совпадают (нашли)
                {
                    //---==StreamId
                    //перевыделим память
                    if (StreamIds == null) StreamIds = new uint[1];
                    else Array.Resize(ref StreamIds, StreamIds.Length + 1);
                    //сохраним найденный Id
                    StreamIds[StreamIds.Length - 1] = Id;

                    //---==Path
                    //перевыделим память
                    if (Paths == null) Paths = new string[1];
                    else Array.Resize(ref Paths, Paths.Length + 1);
                    //сохраним найденный Path
                    Paths[Paths.Length - 1] = curPath + "\\" + curName;
                }

                //идем по Child
                string newPath = curPath + "\\" + curName;
                findInDEArray(DEArray[Id].Child, newPath, Name, ref Paths, ref StreamIds);

                //идем по Left
                findInDEArray(DEArray[Id].LeftSibling, curPath, Name, ref Paths, ref StreamIds);

                //идем по Right
                findInDEArray(DEArray[Id].RightSibling, curPath, Name, ref Paths, ref StreamIds);
            }

            private bool findPathId(uint Id, string Path, out uint foundId)    //поиск Stream Id для заданного пути (вернет true, если найдет)
            {
                //если попали в NOSTREAM
                if (Id == SpecialValues.NOSTREAM)
                {
                    foundId = 0;
                    return false;
                }

                string curName = DEArray[Id].Name.Substring(0, DEArray[Id].Name.IndexOf('\0')); //имя текущей записи
                curName = curName.ToUpper();

                int pathPos = Path.IndexOf('\\');   //позиция первого \ в искомом пути
                //определим имя до первого \ в искомом пути
                string curPath = "";
                string nextPath = "";
                if (pathPos != -1)
                {
                    curPath = Path.Substring(0, pathPos);
                    //nextPath = Path.Substring(pathPos + 1);
                }
                else
                {
                    curPath = Path;
                    //nextPath = Path;
                }

                //проверим текущую запись
                if (curName.CompareTo(curPath) == 0)    //если она подходит
                {
                    if (pathPos == -1)  //текущая запись - та, что мы ищем
                    {
                        foundId = Id;
                        return true;
                    }
                    else    //текущая запись - еще не конец поисков
                    {
                        nextPath = Path.Substring(pathPos + 1);
                        //попробуем пройти к ее ребенку
                        return findPathId(DEArray[Id].Child, nextPath, out foundId);
                    }
                }

                nextPath = Path;

                //не нашли пока - идем по дереву
                uint nextId = 0;
                if (curPath.Length < DEArray[Id].NameLength / 2) nextId = DEArray[Id].LeftSibling;          //если искомое имя короче текущего, то пойдем влево
                else if (curPath.Length > DEArray[Id].NameLength / 2) nextId = DEArray[Id].RightSibling;    //если оно длинее, то пойдем вправо
                else if (curPath.CompareTo(curName) < 0) nextId = DEArray[Id].LeftSibling;              //если искомое имя меньше, то пойдем влево
                else nextId = DEArray[Id].RightSibling;                                                 //если никакое условие выше не сработало, пойдем вправо

                return findPathId(nextId, nextPath, out foundId);   //пошли, куда выбрали
            }
            #endregion

            #region protected internal
            protected internal void showFolderTree()    //отобразить дерево папок
            {
                //выделили память и заполнили данные по Root Entry и ее Child
                FolderTreeEntry[] FTE = new FolderTreeEntry[2];
                FTE[0].TreeLevel = 0;
                FTE[0].Name = DEArray[0].Name.Substring(0, DEArray[0].Name.IndexOf('\0'));
                FTE[0].Name += " <root storage>";
                FTE[1].Parent = FTE[0].Name;
                FTE[1].TreeLevel = 1;

                //идем по дереву
                buildFolderTree(DEArray[0].Child, ref FTE);

                //---==вывод дерева папок на экран
                //отобразим заголовок с именем открытого файла
                FileStream fs = (FileStream)fileReader.BaseStream;
                Console.Clear();
                Console.WriteLine("Folder Tree of " + fs.Name);
                Console.WriteLine();

                //отобразим записи
                Console.WriteLine(FTE[0].Name);
                for (int i = 1; i < FTE.Length; i++)
                {
                    //---==псевдографику нарисуем
                    for (int l = 1; l < FTE[i].TreeLevel; l++)
                    {
                        int j = 0;
                        for (j = i + 1; j < FTE.Length; j++)
                        {
                            if (FTE[j].TreeLevel == l)
                            {
                                Console.Write("│ ");
                                break;
                            }
                        }
                        if (j == FTE.Length) Console.Write("  ");
                    }

                    bool hasSiblingsFurther = false;
                    for (int j = i + 1; j < FTE.Length; j++)
                        if (FTE[j].Parent.CompareTo(FTE[i].Parent) == 0)
                        {
                            hasSiblingsFurther = true;
                            break;
                        }
                    if (hasSiblingsFurther) Console.Write("├─");
                    else Console.Write("└─");

                    //отобразим запись
                    Console.WriteLine(FTE[i].Name);
                }

                Console.WriteLine();
            }

            protected internal bool findStream(string Name, ref string[] Paths, ref uint[] StreamIds)    //поиск потока в файле по заданному имени Name
                                                                                                         //в Paths положит массив путей до найденных потоков (или null, если ничего не нашел)
                                                                                                         //в StreamIds положит StreamId найденных потоков, или null, если ничего не нашел
                                                                                                         //вернет true, если поток найден, или false, если нет
            {
                findInDEArray(0, "", Name, ref Paths, ref StreamIds);   //ищем

                if (Paths == null || StreamIds == null) //если поиск неудачен
                {
                    Paths = null;
                    StreamIds = null;
                    return false;
                }

                if (Paths.Length != StreamIds.Length)  //если результат поиска неадекватен
                {
                    Paths = null;
                    StreamIds = null;
                    return false;
                }

                return true;    //если поиск удачен
            }

            protected internal bool findStream(string Name, ref string[] Paths)    //поиск потока в файле по заданному имени Name
                                                                                   //в Paths положит массив путей до найденных потоков (или null, если ничего не нашел)
                                                                                   //вернет true, если поток найден, или false, если нет
            {
                uint[] StreamIds = null;

                findInDEArray(0, "", Name, ref Paths, ref StreamIds);   //ищем

                if (Paths == null || StreamIds == null) //если поиск неудачен
                {
                    Paths = null;
                    StreamIds = null;
                    return false;
                }

                if (Paths.Length != StreamIds.Length)  //если результат поиска неадекватен
                {
                    Paths = null;
                    StreamIds = null;
                    return false;
                }

                return true;    //если поиск удачен
            }

            protected internal bool getPathId(string Path, out uint Id)  //определить Stream Id в Directory Entry Array для заданного пути (вернет true, если успешно)
            {
                return findPathId(0, Path.ToUpper(), out Id);
            }

            protected internal MemoryStream getStream(uint Id)  //получить из файла поток для чтения по его Stream ID (вернет null, если такого потока нет)
            {
                if (DEArray[Id].ObjectType == 0x00 || DEArray[Id].ObjectType == 0x01) return null;  //если заданный объект имеет тип unknown или storage, вернем null
                
                byte[] byteStream = null;   //поток в виде байтового массива
                MemoryStream memStream;     //поток в виде MemoryStream
                
                switch (DEArray[Id].ObjectType)  //все зависит от того, какого типа объект задан
                {
                    case 0x00:  //type: unknown OR unallocated
                        return null;
                    case 0x01:  //type: storage
                        return null;
                    case 0x02:  //type: stream
                        //найдем размер текущего потока
                        ulong streamSize = (CFHeader.MajorVersion.SequenceEqual(new byte[] { 0x04, 0x00 })) ? DEArray[Id].StreamSizeV4 : DEArray[Id].StreamSizeV3;
                        byteStream = new byte[streamSize];  //выделили память под данные

                        //---==читаем ИЗ ФАЙЛА
                        //определим интересующий нас поток В ФАЙЛЕ
                        //(если размер потока больше отсечки для miniStream, то читать данные будем из FAT секторов,
                        // если он меньше, то из файла прочитаем miniStream, а уже из miniStream потом прочитаем данные)
                        uint readId = (streamSize >= CFHeader.MiniStreamCutoffSize) ? Id : 0;
                        //найдем первый сектор В ФАЙЛЕ
                        uint curSector = DEArray[readId].StartSectorLoc;
                        uint sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.SectorShift, 0));    //размер сектора в файле
                        //найдем размер потока, который будем читать из файла
                        ulong fileBufferSize = (CFHeader.MajorVersion.SequenceEqual(new byte[] { 0x04, 0x00 })) ? DEArray[readId].StreamSizeV4 : DEArray[readId].StreamSizeV3;
                        byte[] fileBuffer = new byte[fileBufferSize];   //считанные из файла данные
                        byte[] readSector = new byte[sectorSize];   //текущий прочитанный из файла сектор
                        int posInArray = 0;    //текущее положение в массиве fileBuffer
                        //читаем из файла
                        while (curSector != SpecialValues.ENDOFCHAIN)   //пока не достигнем конца цепочки секторов
                        {
                            uint fileOffset = (curSector + 1) * sectorSize; //положение текущего сектора в файле
                            fileReader.BaseStream.Seek(fileOffset, SeekOrigin.Begin);   //нашли нужный сектор в файле
                            readSector = fileReader.ReadBytes((int)sectorSize); //считали сектор из файла
                            curSector = FAT[curSector];     //нашли следующий сектор
                            if (curSector == SpecialValues.ENDOFCHAIN) //если текущий сектор последний
                            {
                                int numLastBytes = (int)Math.IEEERemainder(fileBufferSize, sectorSize); //кол-во использованных под поток байт в последнем секторе
                                readSector.Take(numLastBytes).ToArray().CopyTo(fileBuffer, posInArray); //перенесли только использованные байты в fileBuffer
                                posInArray += numLastBytes; //увеличили счетчик положения в fileBuffer
                            }
                            else                                        //если текущий сектор не последний
                            {
                                readSector.CopyTo(fileBuffer, posInArray);  //перенесли весь сектор в fileBuffer
                                posInArray += (int)sectorSize;  //увеличили счетчик положения в fileBuffer

                            }
                        }

                        if (streamSize >= CFHeader.MiniStreamCutoffSize)   //если размер потока больше отсечки для miniStream
                        {
                            //то прочитанные в fileBuffer данные - те, которые были запрошены
                            fileBuffer.CopyTo(byteStream, 0);   //просто перенесем их в поток
                        }
                        else                                                //если размер потока меньше отсечки для miniStream
                        {
                            //то из файла мы прочитали только miniStream
                            //а теперь из miniStream прочитаем запрощенные данные
                            MemoryStream miniStream = new MemoryStream(fileBuffer); //создали из прочитанных из файла данных поток в памяти
                            BinaryReader brMiniStream = new BinaryReader(miniStream);   //создали читальщик двоичных данных из нового потока

                            curSector = DEArray[Id].StartSectorLoc; //первый сектор с запроошенными данными в miniStream 
                            sectorSize = (uint)Math.Pow(2, BitConverter.ToUInt16(CFHeader.MiniSectorShift, 0));    //размер сектора в miniStream
                            Array.Resize(ref readSector, (int)sectorSize);  //перевыделили память под текущий считанный из miniStream сектор
                            posInArray = 0; //позиция в byteStream;
                            //читаем из miniStream
                            while (curSector != SpecialValues.ENDOFCHAIN)   //пока не достигнем конца цепочки секторов
                            {
                                uint msOffset = (curSector + 1) * sectorSize; //положение текущего сектора в miniStream
                                miniStream.Seek(msOffset, SeekOrigin.Begin);   //нашли нужный сектор в miniStream
                                readSector = brMiniStream.ReadBytes((int)sectorSize); //считали сектор из miniStream
                                curSector = miniFAT[curSector];     //нашли следующий сектор
                                if (curSector == SpecialValues.ENDOFCHAIN) //если текущий сектор последний
                                {
                                    int numLastBytes = (int)Math.IEEERemainder(streamSize, sectorSize); //кол-во использованных под поток байт в последнем секторе
                                    readSector.Take(numLastBytes).ToArray().CopyTo(byteStream, posInArray); //перенесли только использованные байты в byteStream
                                    posInArray += numLastBytes; //увеличили счетчик положения в byteStream
                                }
                                else                                        //если текущий сектор не последний
                                {
                                    readSector.CopyTo(byteStream, posInArray);  //перенесли весь сектор в byteStream
                                    posInArray += (int)sectorSize;  //увеличили счетчик положения в byteStream

                                }
                            }
                        }
                        break;
                    case 0x05:  //type: root storage                        
                        return null;
                    default:    //unknown type
                        return null;
                }

                //создадим memStream из byteStream и вернем его из метода
                memStream = new MemoryStream(byteStream);
                return memStream;
            }

            protected internal MemoryStream getStream(string Path)  //получить из файла поток для чтения по его пути (вернет null, если такого потока нет)
            {
                uint Id;
                if (getPathId(Path, out Id)) return getStream(Id);  //если нашли заданный поток, вернем его MemoryStream
                return null;    //если не нашли, вернем null
            }
            #endregion
            #endregion
        }
        #endregion

        #region Структуры
        private struct FileInformationBlock //the File Information Block structure
        {
            internal FibBase fibBase;           //FibBase structure [32 bytes]
            internal ushort csw;                //Count of 16-bit values corresponding to fibRgW that follow (MUST: 0x000E) [2 bytes]
            internal FibRgW97 fibRgW;           //FibRgW97 structure [28 bytes]
            internal ushort cslw;               //Count of 32-bit values corresponding to fibRgLw that follow (MUST: 0x0016) [2 bytes]
            internal FibRgLw97 fibRgLw;         //FibRgLw97 structure [88 bytes]
            internal ushort cbRgFcLcb;          //Count of 64-bit values corresponding to fibRgFcLcbBlob that follow
                                                //(MUST:
                                                //0x005D if nFib=0x00C1 OR
                                                //0x006C if nFib=0x00D9 OR
                                                //0x0088 if nFib=0x0101 OR
                                                //0x00A4 if nFib=0x010C OR
                                                //0x00B7 if nFib=0x0112)
                                                //[2 bytes]
            internal FibRgFcLcb fibRgFcLcbBlob; //FibRgFcLcb structure [variable]
        }

        private struct FibBase  //FibBase structure
        {
            internal ushort wIdent;     //Specifies that this is a Word Binary File (MUST: 0xA5EC) [2 bytes]
            internal ushort nFib;       //Version number of the file format used. Superseded by FibRgCswNew.nFibNew if it is present (SHOULD: 0x00C0, 0x00C1 OR 0x00C2) [2 bytes]
            internal byte[] unused;     //Undefined value (MUST: be ignored) [2 bytes]
            internal ushort lid;        //Two Digits Hexadecimal Language Code [2 bytes]
            internal ushort pnNext;     //Offset in WordDocument stream of the FIB for the document which contains all the AutoText items. [2 bytes]
            internal BitArray bitsAtoM; //Bit flags from A to M [2 bytes]
            internal ushort nFibBack;   //(MUST: 0x00BF OR 0x00C1) [2 bytes]
            internal uint lKey;         //If fEncrypted=1 and fObfuscation=1, specifies the XOR obfuscation password verifier.
                                        //If fEncrypted=1 and fObfuscation=0, specifies the size of the EncryptionHeader stored at the beginning of the Table stream
                                        //(MUST be 0 otherwise) [4 bytes]
            internal byte envr;         //(MUST be 0, MUST be ignored) [1 byte]
            internal BitArray bitsNtoS; //Bit flags from N to S [1 byte]
            internal ushort reserved3;  //(MUST be 0, MUST be ignored) [2 bytes]
            internal ushort reserved4;  //(MUST be 0, MUST be ignored) [2 bytes]
            internal uint reserved5;    //(MUST be ignored) [4 bytes]
            internal uint reserved6;    //(MUST be ignored) [4 bytes]

            //NOTE: FOLLOWING FIELDS ARE NOT IN THE DOCUMENT!!! THEY ARE JUST FOR CONVENIENCE
            //---== bitsAtoM decomposed ==---
            internal bool fDot;                 //A [1 bit]: Specifies whether this document is a document template
            internal bool fGlsy;                //B [1 bit]: Specifies whether this document contains only AutoText items
            internal bool fComplex;             //C [1 bit]: Specifies that the last save operation was an incremental save operation
            internal bool fHasPic;              //D [1 bit]: When set to 0, there SHOULD be no pictures in the document
            internal byte cQuickSaves;          //E [4 bits]: Number of consecutive times document was incrementally saved (MUST: 0xF, if nFib>=0x00D9)
            internal bool fEncrypted;           //F [1 bit]: Specifies whether the document is encrypted or obfuscated
            internal bool fWhichTblStm;         //G [1 bit]: Specifies the Table stream to which the FIB refers. When true, use 1Table; when false, use 0Table
            internal bool fReadOnlyRecommended; //H [1 bit]: Specifies that author recommended this document to be opened in read-only mode
            internal bool fWriteReservation;    //I [1 bit]: Specifies whether the document has a write-reservation password
            internal bool fExtChar;             //J [1 bit]: MUST be true
            internal bool fLoadOverride;        //K [1 bit]: Specifies whether to override the language information and font that are specified in the paragraph style
                                                //at istd 0 (the normal style) with the defaults
            internal bool fFarEast;             //L [1 bit]: Specifies whether the language of the application that created the document was an East Asian language
            internal bool fObfuscated;          //M [1 bit]: If fEncrypted is 1, this bit specifies whether the document is obfuscated by using XOR obfuscation; MUST be ignored otherwise
            //---== end of decomposition ==---

            //---== bitsNtoS decomposed ==---
            internal bool fMac;                 //N [1 bit]: (MUST be 0, MUST be ignored)
            internal bool fEmptySpecial;        //O [1 bit]: (SHOULD be 0, SHOULD be ignored)
            internal bool fLoadOverridePage;    //P [1 bit]: Specifies whether to override the section properties for page size, orientation, and margins with the defaults
            internal bool reserved1;            //Q [1 bit]: (MUST be ignored)
            internal bool reserved2;            //R [1 bit]: (MUST be ignored)
            internal byte fSpare0;              //S [1 bit]: (MUST be ignored)
            //---== end of decomposition ==---
        }

        private struct FibRgW97 //FibRgW97 structure
        {
            internal ushort reserved1;  //(MUST be ignored) [2 bytes]
            internal ushort reserved2;  //(MUST be ignored) [2 bytes]
            internal ushort reserved3;  //(MUST be ignored) [2 bytes]
            internal ushort reserved4;  //(MUST be ignored) [2 bytes]
            internal ushort reserved5;  //(SHOULD: 0, MUST be ignored) [2 bytes]
            internal ushort reserved6;  //(SHOULD: 0, MUST be ignored) [2 bytes]
            internal ushort reserved7;  //(SHOULD: 0, MUST be ignored) [2 bytes]
            internal ushort reserved8;  //(SHOULD: 0, MUST be ignored) [2 bytes]
            internal ushort reserved9;  //(SHOULD: 0, MUST be ignored) [2 bytes]
            internal ushort reserved10; //(SHOULD: 0, MUST be ignored) [2 bytes]
            internal ushort reserved11; //(SHOULD: 0, MUST be ignored) [2 bytes]
            internal ushort reserved12; //(SHOULD: 0, MUST be ignored) [2 bytes]
            internal ushort reserved13; //(SHOULD: 0, MUST be ignored) [2 bytes]
            internal ushort lidFE;      //If nFib=0x00C1: if FibBase.fFarEast=true, this is the LID of the stored style names (MUST be ignored otherwise)
                                        //If nFib=0x00D9 OR 0x0101 OR 0x010C OR 0x0112, this is the LID of the stored style names
                                        //[2 bytes]
        }

        private struct FibRgLw97    //FibRgLw97 structure
        {
            internal uint cbMac;        //Specifies the count of bytes in WordDocument stream that has any meaning. All bytes in WordDocument at offset cbMac and greater MUST be ignored [4 bytes]
            internal uint reserved1;    //(MUST be ignored) [4 bytes]
            internal uint reserved2;    //(MUST be ignored) [4 bytes]
            internal uint ccpText;      //Count of CPs in the Main Document (MUST: >=0) [4 bytes]
            internal uint ccpFtn;       //Count of CPs in the Footnote Subdocument (MUST: >=0) [4 bytes]
            internal uint ccpHdd;       //Count of CPs in the Header Subdocument (MUST: >=0) [4 bytes]
            internal uint reserved3;    //(MUST be ignored) [4 bytes]
            internal uint ccpAtn;       //Count of CPs in the Comment Subdocument (MUST: >=0) [4 bytes]
            internal uint ccpEdn;       //Count of CPs in the Endnote Subdocument (MUST: >=0) [4 bytes]
            internal uint ccpTxbx;      //Count of CPs in the Textbox Subdocument of the Main Document (MUST: >=0) [4 bytes]
            internal uint ccpHdrTxbx;   //Count of CPs in the Textbox Subdocument of the Header (MUST: >=0) [4 bytes]
            internal uint reserved4;    //(MUST be ignored) [4 bytes]
            internal uint reserved5;    //(MUST be ignored) [4 bytes]
            internal uint reserved6;    //(MUST be ignored) [4 bytes]
            internal uint reserved7;    //(MUST be ignored) [4 bytes]
            internal uint reserved8;    //(MUST be ignored) [4 bytes]
            internal uint reserved9;    //(MUST be ignored) [4 bytes]
            internal uint reserved10;   //(MUST be ignored) [4 bytes]
            internal uint reserved11;   //(MUST be ignored) [4 bytes]
            internal uint reserved12;   //(MUST be ignored) [4 bytes]
            internal uint reserved13;   //(MUST be ignored) [4 bytes]
            internal uint reserved14;   //(MUST be ignored) [4 bytes]
        }

        private struct FibRgFcLcb   //FibRgFcLcb structure
        {
            //NOTE: Containment depends of nFib value:
            //nFib=0x00C1:  use only v97 portion [744 bytes]
            //nFib=0x00D9:  use only v97 and v00 portions [864 bytes]
            //nFib=0x0101:  use only v97, v00 and v02 portions [1088 bytes]
            //nFib=0x010C:  use only v97, v00, v02 and v03 portions [1312 bytes]
            //nFib=0x0112:  use all (v97, v00, v02, v03 and v07) portions [1464 bytes]

            internal FibRgFcLcb97 v97;      //FibRgFcLcb97 structure [744 bytes]
            internal FibRgFcLcb2000 v00;    //FibRgFcLcb2000 portion structure [120 bytes]
            internal FibRgFcLcb2002 v02;    //FibRgFcLcb2002 portion structure [224 bytes]
            internal FibRgFcLcb2003 v03;    //FibRgFcLcb2003 portion structure [224 bytes]
            internal FibRgFcLcb2007 v07;    //FibRgFcLcb2007 portion structure [152 bytes]
        }

        private struct FibRgFcLcb97 //FibRgFcLcb97 structure
        {
            internal uint fcStshfOrig;          //(MUST be ignored) [4 bytes]
            internal uint lcbStshfOrig;         //(MUST be ignored) [4 bytes]
            internal uint fcStshf;              //Offset of the STSH in the Table Stream [4 bytes]
            internal uint lcbStshf;             //Size, in bytes, of the STSH (MUST: non zero) [4 bytes]
            internal uint fcPlcffndRef;         //Offset of the PlcffndRef in the Table Stream (MUST be ignored if lcbPlcffndRef=0) [4 bytes]
            internal uint lcbPlcffndRef;        //Size, in bytes, of the PlcffndRef [4 bytes]
            internal uint fcPlcffndTxt;         //Offset of the PlcffndTxt in the Table Stream (MUST be ignored if lcbPlcffndTxt=0) [4 bytes]
            internal uint lcbPlcffndTxt;        //Size, in bytes, of the PlcffndTxt (MUST: 0 if FibRgLw97.ccpFtn=0 AND nonzero if FibRgLw97.ccpFtn!=0) [4 bytes]
            internal uint fcPlcfandRef;         //Offset of the PlcfandRef in the Table Stream (MUST be ignored if lcbPlcfandRef=0) [4 bytes]
            internal uint lcbPlcfandRef;        //Size, in bytes, of the PlcfandRef
            internal uint fcPlcfandTxt;         //Offset of the PlcfandTxt in the Table Stream (MUST be ignored if lcbPlcfandTxt=0) [4 bytes]
            internal uint lcbPlcfandTxt;        //Size, in bytes, of the PlcfandTxt (MUST: 0 if FibRgLw97.ccpAtn=0 AND nonzero if FibRgLw97.ccpAtn!=0) [4 bytes]
            internal uint fcPlcfSed;            //Offset of the PlcfSed in the Table Stream (MUST be ignored if lcbPlcfSed=0) [4 bytes]
            internal uint lcbPlcfSed;           //Size, in bytes, of the PlcfSed [4 bytes]
            internal uint fcPlcPad;             //(MUST be ignored) [4 bytes]
            internal uint lcbPlcPad;            //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcPlcfPhe;            //Offset of the Plc in the Table Stream (this Plc SHOULD be ignored) [4 bytes]
            internal uint lcbPlcfPhe;           //Size, in bytes, of the Plc at offset fcPlcfPhe in the Table Stream [4 bytes]
            internal uint fcSttbfGlsy;          //Offset of the SttbfGlsy in the Table Stream [4 bytes]
            internal uint lcbSttbfGlsy;         //Size, in bytes, of the SttbfGlsy (MUST: 0 if FibBase.fGlsy=0) [4 bytes]
            internal uint fcPlcfGlsy;           //Offset of the PlcfGlsy in the Table Stream [4 bytes]
            internal uint lcbPlcfGlsy;          //Size, in bytes, of the PlcfGlsy (MUST: 0 if FibBase.fGlsy=0) [4 bytes]
            internal uint fcPlcfHdd;            //Offset of the Plcfhdd in the Table Stream (MUST be ignored if lcbPlcfHdd=0) [4 bytes]
            internal uint lcbPlcfHdd;           //Size, in bytes, of the Plcfhdd (MUST: 0 if there is no Plcfhdd) [4 bytes]
            internal uint fcPlcfBteChpx;        //Offset of the PlcBteChpx in the Table Stream (MUST: >0 AND be a valid offset in Table Stream) [4 bytes]
            internal uint lcbPlcfBteChpx;       //Size, in bytes, of the PlcBteChpx (MUST: >0) [4 bytes]
            internal uint fcPlcfBtePapx;        //Offset of the PlcBtePapx in the Table Stream (MUST: >0 AND be a valid offset in Table Stream) [4 bytes]
            internal uint lcbPlcfBtePapx;       //Size, in bytes, of the PlcBtePapx (MUST: >0) [4 bytes]
            internal uint fcPlcfSea;            //(MUST be ignored) [4 bytes]
            internal uint lcbPlcfSea;           //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcSttbfFfn;           //Offset of the SttbfFfn in the Table Stream (MUST be ignored if lcbSttbfFfn=0) [4 bytes]
            internal uint lcbSttbfFfn;          //Size, in bytes, of the SttbfFfn [4 bytes]
            internal uint fcPlcfFldMom;         //Offset of the PlcFld for Main Document in the Table Stream (MUST be ignored if lcbPlcfFldMom=0) [4 bytes]
            internal uint lcbPlcfFldMom;        //Size, in bytes, of the PlcFld for Main Document [4 bytes]
            internal uint fcPlcfFldHdr;         //Offset of the PlcFld for Header Document in the Table Stream (MUST be ignored if lcbPlcfFldHdr=0) [4 bytes]
            internal uint lcbPlcfFldHdr;        //Size, in bytes, of the PlcFld for Header Document [4 bytes]
            internal uint fcPlcfFldFtn;         //Offset of the PlcFld for Footnote Document in the Table Stream (MUST be ignored if lcbPlcfFldFtn=0) [4 bytes]
            internal uint lcbPlcfFldFtn;        //Size, in bytes, of the PlcFld for Footnote Document [4 bytes]
            internal uint fcPlcfFldAtn;         //Offset of the PlcFld for Comment Document in the Table Stream (MUST be ignored if lcbPlcfFldAtn=0) [4 bytes]
            internal uint lcbPlcfFldAtn;        //Size, in bytes, of the PlcFld for Comment Document [4 bytes]
            internal uint fcPlcfFldMcr;         //(MUST be ignored) [4 bytes]
            internal uint lcbPlcfFldMcr;        //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcSttbfBkmk;          //Offset of the SttbfBkmk in the Table Stream (MUST be ignored if lcbSttbfBkmk=0) [4 bytes]
            internal uint lcbSttbfBkmk;         //Size, in bytes, of the SttbfBkmk [4 bytes]
            internal uint fcPlcfBkf;            //Offset of the PlcfBkf in the Table Stream (MUST be ignored if lcbPlcfBkf=0) [4 bytes]
            internal uint lcbPlcfBkf;           //Size, in bytes, of the PlcfBkf [4 bytes]
            internal uint fcPlcfBkl;            //Offset of the PlcfBkl in the Table Stream (MUST be ignored if lcbPlcfBkl=0) [4 bytes]
            internal uint lcbPlcfBkl;           //Size, in bytes, of the PlcfBkl [4 bytes]
            internal uint fcCmds;               //Offset of the Tcg in the Table Stream (MUST be ignored if lcbCmds=0) [4 bytes]
            internal uint lcbCmds;              //Size, in bytes, of the Tcg [4 bytes]
            internal uint fcUnused1;            //(MUST be ignored) [4 bytes]
            internal uint lcbUnused1;           //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcSttbfMcr;           //(MUST be ignored) [4 bytes]
            internal uint lcbSttbfMcr;          //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcPrDrvr;             //Offset of the PrDrvr in the Table Stream (MUST be ignored if lcbPrDrvr=0) [4 bytes]
            internal uint lcbPrDrvr;            //Size, in bytes, of the PrDrvr [4 bytes]
            internal uint fcPrEnvPort;          //Offset of the PrEnvPort in the Table Stream (MUST be ignored if lcbPrEnvPort=0) [4 bytes]
            internal uint lcbPrEnvPort;         //Size, in bytes, of the PrEnvPort [4 bytes]
            internal uint fcPrEnvLand;          //Offset of the PrEnvLand in the Table Stream (MUST be ignored if lcbPrEnvLand=0) [4 bytes]
            internal uint lcbPrEnvLand;         //Size, in bytes, of the PrEnvLand [4 bytes]
            internal uint fcWss;                //Offset of the Selsf in the Table Stream (MUST be ignored if lcbWss=0) [4 bytes]
            internal uint lcbWss;               //Size, in bytes, of the Selsf [4 bytes]
            internal uint fcDop;                //Offset of the Dop in the Table Stream [4 bytes]
            internal uint lcbDop;               //Size, in bytes, of the Dop (MUST: be nonzero) [4 bytes]
            internal uint fcSttbfAssoc;         //Offset of the SttbfAssoc in the Table Stream [4 bytes]
            internal uint lcbSttbfAssoc;        //Size, in bytes, of the SttbfAssoc (MUST: be nonzero) [4 bytes]
            internal uint fcClx;                //Offset of the Clx in the Table Stream [4 bytes]
            internal uint lcbClx;               //Size, in bytes, of the Clx (MUST: >0) [4 bytes]
            internal uint fcPlcfPgdFtn;         //(MUST be ignored) [4 bytes]
            internal uint lcbPlcfPgdFtn;        //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcAutosaveSource;     //(MUST be ignored) [4 bytes]
            internal uint lcbAutosaveSource;    //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcGrpXstAtnOwners;    //Offset of the array of XSTs in the Table Stream [4 bytes]
            internal uint lcbGrpXstAtnOwners;   //Size, in bytes, of the XST array [4 bytes]
            internal uint fcSttbfAtnBkmk;       //Offset of the SttbfAtnBkmk in the Table Stream (MUST be ignored if lcbSttbfAtnBkmk=0) [4 bytes]
            internal uint lcbSttbfAtnBkmk;      //Size, in bytes, of the SttbfAtnBkmk [4 bytes]
            internal uint fcUnused2;            //(MUST be ignored) [4 bytes]
            internal uint lcbUnused2;           //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcUnused3;            //(MUST be ignored) [4 bytes]
            internal uint lcbUnused3;           //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcPlcSpaMom;          //Offset of the PlcfSpa for Main Document in the Table Stream [4 bytes]
            internal uint lcbPlcSpaMom;         //Size, in bytes, of the PlcfSpa for Main Document [4 bytes]
            internal uint fcPlcSpaHdr;          //Offset of the PlcfSpa for the Header Document in the Table Stream [4 bytes]
            internal uint lcbPlcSpaHdr;         //Size, in bytes, of the PlcfSpa for the Header Document [4 bytes]
            internal uint fcPlcfAtnBkf;         //Offset of the PlcfBkf in the Table Stream (MUST be ignored if lcbPlcfAtnBkf=0) [4 bytes]
            internal uint lcbPlcfAtnBkf;        //Size, in bytes, of the PlcfBkf [4 bytes]
            internal uint fcPlcfAtnBkl;         //Offset of the PlcfBkl in the Table Stream (MUST be ignored if lcbPlcfAtnBkl=0) [4 bytes]
            internal uint lcbPlcfAtnBkl;        //Size, in bytes, of the PlcfBkl [4 bytes]
            internal uint fcPms;                //Offset of the Pms in the Table Stream (MUST be ignored if lcbPms=0) [4 bytes]
            internal uint lcbPms;               //Size, in bytes, of the Pms [4 bytes]
            internal uint fcFormFldSttbs;       //(MUST be ignored) [4 bytes]
            internal uint lcbFormFldSttbs;      //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcPlcfendRef;         //Offset of the PlcfendRef in the Table Stream (MUST be ignored if lcbPlcfendRef=0) [4 bytes]
            internal uint lcbPlcfendRef;        //Size, in bytes, of the PlcfendRef [4 bytes]
            internal uint fcPlcfendTxt;         //Offset of the PlcfendTxt in the Table Stream (MUST be ignored if lcbPlcfendTxt=0) [4 bytes]
            internal uint lcbPlcfendTxt;        //Size, in bytes, of the PlcfendTxt (MUST: 0 if FibRgLw97.ccpEdn=0 AND nonzero if FibRgLw97.ccpEdn is nonzero) [4 bytes]
            internal uint fcPlcfFldEdn;         //Offset of the PlcFld for the Endnote Document in the Table Stream (MUST be ignored if lcbPlcfFldEdn=0) [4 bytes]
            internal uint lcbPlcfFldEdn;        //Size, in bytes, of the PlcFld for the Endnote Document [4 bytes]
            internal uint fcUnused4;            //(MUST be ignored) [4 bytes]
            internal uint lcbUnused4;           //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcDggInfo;            //Offset of the OfficeArtContent in the Table Stream [4 bytes]
            internal uint lcbDggInfo;           //Size, in bytes, of the OfficeArtContent [4 bytes]
            internal uint fcSttbfRMark;         //Offset of the SttbfRMark in the Table Stream (MUST be ignored if lcbSttbfRMark=0) [4 bytes]
            internal uint lcbSttbfRMark;        //Size, in bytes, of the SttbfRMark [4 bytes]
            internal uint fcSttbfCaption;       //Offset of the SttbfCaption in the Table Stream (MUST be ignored if lcbSttbfCaption=0 OR if this document is not the Normal template) [4 bytes]
            internal uint lcbSttbfCaption;      //Size, in bytes, of the SttbfCaption (MUST: 0 if FibBase.fDot=0) [4 bytes]
            internal uint fcSttbfAutoCaption;   //Offset of the SttbfAutoCaption in the Table Stream (MUST be ignored if lcbSttbfAutoCaption=0 OR if this document is not the Normal template) [4 bytes]
            internal uint lcbSttbfAutoCaption;  //Size, in bytes, of the SttbfAutoCaption (MUST: 0 if FibBase.fDot=0) [4 bytes]
            internal uint fcPlcfWkb;            //Offset of the PlcfWKB in the Table Stream (MUST be ignored if lcbPlcfWkb=0) [4 bytes]
            internal uint lcbPlcfWkb;           //Size, in bytes, of the PlcfWKB [4 bytes]
            internal uint fcPlcfSpl;            //Offset of the Plcfspl in the Table Stream (MUST be ignored if lcbPlcfSpl=0) [4 bytes]
            internal uint lcbPlcfSpl;           //Size, in bytes, of the Plcfspl [4 bytes]
            internal uint fcPlcftxbxTxt;        //Offset of the PlcftxbxTxt in the Table Stream (MUST be ignored if lcbPlcftxbxTxt=0) [4 bytes]
            internal uint lcbPlcftxbxTxt;       //Size, in bytes, of the PlcftxbxTxt (MUST: 0 if FibRgLw97.ccpTxbx=0 AND nonzero if FibRgLw97.ccpTxbx is nonzero) [4 bytes]
            internal uint fcPlcfFldTxbx;        //Offset of the PlcFld for the Textbox Document in the Table Stream (MUST be ignored if lcbPlcfFldTxbx=0) [4 bytes]
            internal uint lcbPlcfFldTxbx;       //Size, in bytes, of the PlcFld for the Textbox Document [4 bytes]
            internal uint fcPlcfHdrtxbxTxt;     //Offset of the PlcfHdrtxbxTxt in the Table Stream [4 bytes]
            internal uint lcbPlcfHdrtxbxTxt;    //Size, in bytes, of the PlcfHdrtxbxTxt (MUST: 0 if FibRgLw97.ccpHdrTxbx=0 AND nonzero if FibRgLw97.ccpHdrTxbx is nonzero) [4 bytes]
            internal uint fcPlcffldHdrTxbx;     //Offset of the PlcFld for the Header Textbox in the Table Stream (MUST be ignored if lcbPlcffldHdrTxbx=0) [4 bytes]
            internal uint lcbPlcffldHdrTxbx;    //Size, in bytes, of the PlcFld for the Header Textbox [4 bytes]
            internal uint fcStwUser;            //Offset of the StwUser into the Table Stream (MUST be ignored if lcbStwUser=0) [4 bytes]
            internal uint lcbStwUser;           //Size, in bytes, of the StwUser at offset [4 bytes]
            internal uint fcSttbTtmbd;          //Offset of the SttbTtmbd into the Table Stream (MUST be ignored if lcbSttbTtmbd=0) [4 bytes]
            internal uint lcbSttbTtmbd;         //Size, in bytes, of the SttbTtmbd [4 bytes]
            internal uint fcCookieData;         //Offset of the RgCdb in the Table Stream (MUST be ignored if lcbCookieData=0. MAY be ignored otherwise) [4 bytes]
            internal uint lcbCookieData;        //Size, in bytes, of the RgCdb [4 bytes]
            internal uint fcPgdMotherOldOld;    //Offset of the deprecated document page layout cache in the Table Stream (MUST be ignored if lcbPgdMotherOldOld=0) [4 bytes]
            internal uint lcbPgdMotherOldOld;   //Size, in bytes, of the deprecated document page layout cache [4 bytes]
            internal uint fcBkdMotherOldOld;    //Offset of the Deprecated document text flow break cache in the Table Stream (MUST be ignored if lcbBkdMotherOldOld=0) [4 bytes]
            internal uint lcbBkdMotherOldOld;   //Size, in bytes, of the deprecated document text flow break cache [4 bytes]
            internal uint fcPgdFtnOldOld;       //Offset of the Deprecated footnote layout cache in the Table Stream (MUST be ignored if lcbPgdFtnOldOld=0) [4 bytes]
            internal uint lcbPgdFtnOldOld;      //Size, in bytes, of the deprecated footnote layout cache [4 bytes]
            internal uint fcBkdFtnOldOld;       //Offset of the deprecated footnote text flow break cache in the Table Stream (MUST be ignored if lcbBkdFtnOldOld=0) [4 bytes]
            internal uint lcbBkdFtnOldOld;      //Size, in bytes, of the deprecated footnote text flow break cache [4 bytes]
            internal uint fcPgdEdnOldOld;       //Offset of the deprecated endnote layout cache in the Table Stream (MUST be ignored if lcbPgdEdnOldOld=0) [4 bytes]
            internal uint lcbPgdEdnOldOld;      //Size, in bytes, of the deprecated endnote layout cache [4 bytes]
            internal uint fcBkdEdnOldOld;       //Offset of the deprecated endnote text flow break cache in the Table Stream (MUST be ignored if lcbBkdEdnOldOld=0) [4 bytes]
            internal uint lcbBkdEdnOldOld;      //Size, in bytes, of the deprecated endnote text flow break cache [4 bytes]
            internal uint fcSttbfIntlFld;       //(MUST be ignored) [4 bytes]
            internal uint lcbSttbfIntlFld;      //(MUST: 0 AND be ignored) [4 bytes]
            internal uint fcRouteSlip;          //Offset of the RouteSlip in the Table Stream (SHOULD be ignored) [4 bytes]
            internal uint lcbRouteSlip;         //Size, in bytes, of the RouteSlip [4 bytes]
            internal uint fcSttbSavedBy;        //Offset of the SttbSavedBy in the Table Stream (SHOULD be ignored) [4 bytes]
            internal uint lcbSttbSavedBy;       //Size, in bytes, of the SttbSavedBy (SHOULD: 0) [4 bytes]
            internal uint fcSttbFnm;            //Offset of the SttbFnm in the Table Stream (MUST be ignored if lcbSttbFnm=0) [4 bytes]
            internal uint lcbSttbFnm;           //Size, in bytes, of the SttbFnm [4 bytes]
            internal uint fcPlfLst;             //Offset of the PlfLst in the Table Stream (MUST be ignored if lcbPlfLst=0) [4 bytes]
            internal uint lcbPlfLst;            //Size, in bytes, of the PlfLst [4 bytes]
            internal uint fcPlfLfo;             //Offset of the PlfLfo in the Table Stream (MUST be ignored if lcbPlfLfo=0) [4 bytes]
            internal uint lcbPlfLfo;            //Size, in bytes, of the PlfLfo [4 bytes]
            internal uint fcPlcfTxbxBkd;        //Offset of the PlcftxbxBkd in the Table Stream [4 bytes]
            internal uint lcbPlcfTxbxBkd;       //Size, in bytes, of the PlcftxbxBkd (MUST: 0 if FibRgLw97.ccpTxbx=0 AND nonzero if FibRgLw97.ccpTxbx is nonzero) [4 bytes]
            internal uint fcPlcfTxbxHdrBkd;     //Offset of the PlcfTxbxHdrBkd in the Table Stream [4 bytes]
            internal uint lcbPlcfTxbxHdrBkd;    //Size, in bytes, of the PlcfTxbxHdrBkd (MUST: 0 if FibRgLw97.ccpHdrTxbx=0 AND nonzero if FibRgLw97.ccpHdrTxbx is nonzero) [4 bytes]
            internal uint fcDocUndoWord9;       //Offset of the Version-specific undo information in the WordDocument Stream (SHOULD be ignored) [4 bytes]
            internal uint lcbDocUndoWord9;      //If this is nonzero, version-specific undo information exists at offset fcDocUndoWord9 in the WordDocument Stream [4 bytes]
            internal uint fcRgbUse;             //Offset of the Version-specific undo information in the WordDocument Stream (SHOULD be ignored) [4 bytes]
            internal uint lcbRgbUse;            //Size, in bytes, of the version-specific undo information [4 bytes]
            internal uint fcUsp;                //Offset of the Version-specific undo information in the WordDocument Stream (SHOULD be ignored) [4 bytes]
            internal uint lcbUsp;               //Size, in bytes, of the version-specific undo information [4 bytes]
            internal uint fcUskf;               //Offset of the Version-specific undo information in the Table Stream (SHOULD be ignored) [4 bytes]
            internal uint lcbUskf;              //Size, in bytes, of the version-specific undo information [4 bytes]
            internal uint fcPlcupcRgbUse;       //Offset of the Plc for the version-specific undo information in the Table Stream (SHOULD be ignored) [4 bytes]
            internal uint lcbPlcupcRgbUse;      //Size, in bytes, of the Plc for the version-specific undo information [4 bytes]
            internal uint fcPlcupcUsp;          //Offset of the Plc for the version-specific undo information in the Table Stream (SHOULD be ignored) [4 bytes]
            internal uint lcbPlcupcUsp;         //Size, in bytes, of the Plc for the version-specific undo information [4 bytes]
            internal uint fcSttbGlsyStyle;      //Offset of the SttbGlsyStyle in the Table Stream [4 bytes]
            internal uint lcbSttbGlsyStyle;     //Size, in bytes, of the SttbGlsyStyle (MUST: 0 if FibBase.fGlsy=0) [4 bytes]
            internal uint fcPlgosl;             //Offset of the PlfGosl in the Table Stream (MUST be ignored if lcbPlgosl=0) [4 bytes]
            internal uint lcbPlgosl;            //Size, in bytes, of the PlfGosl [4 bytes]
            internal uint fcPlcocx;             //Offset of the RgxOcxInfo in the Table Stream (MUST:0 AND be ignored if there are no OLE controls in the document) [4 bytes]
            internal uint lcbPlcocx;            //Size, in bytes, of the RgxOcxInfo (MUST:0 AND be ignored if there are no OLE controls in the document) [4 bytes]
            internal uint fcPlcfBteLvc;         //Offset of the deprecated numbering field cache in the Table Stream (MUST be ignored if lcbPlcBteLvc=0. SHOULD be ignored) [4 bytes]
            internal uint lcbPlcfBteLvc;        //Size, in bytes, of the deprecated numbering field cache at offset fcPlcfBteLvc in the Table Stream (SHOULD: 0) [4 bytes]
            internal uint dwLowDateTime;        //The low-order part of a FILETIME structure that specifies when the document was last saved [4 bytes]
            internal uint dwHighDateTime;       //The high-order part of a FILETIME structure that specifies when the document was last saved [4 bytes]
            internal uint fcPlcfLvcPre10;       //Offset of the deprecated list level cache in the Table Stream (MUST be ignored if lcbPlcfLvcPre10=0. SHOULD be ignored) [4 bytes]
            internal uint lcbPlcfLvcPre10;      //Size, in bytes, of the deprecated list level cache at offset fcPlcfLvcPre10 in the Table Stream (SHOULD: 0) [4 bytes]
            internal uint fcPlcfAsumy;          //Offset of the PlcfAsumy in the Table Stream (MUST be ignored if lcbPlcfAsumy=0) [4 bytes]
            internal uint lcbPlcfAsumy;         //Size, in bytes, of the PlcfAsumy [4 bytes]
            internal uint fcPlcfGram;           //Offset of the Plcfgram in the Table Stream (MUST be ignored if lcbPlcfGram=0) [4 bytes]
            internal uint lcbPlcfGram;          //Size, in bytes, of the Plcfgram [4 bytes]
            internal uint fcSttbListNames;      //Offset of the SttbListNames in the Table Stream (MUST be ignored if lcbSttbListNames=0) [4 bytes]
            internal uint lcbSttbListNames;     //Size, in bytes, of the SttbListNames [4 bytes]
            internal uint fcSttbfUssr;          //Offset of the deprecated, version-specific undo information in the Table Stream (SHOULD be ignored) [4 bytes]
            internal uint lcbSttbfUssr;         //Size, in bytes, of the deprecated,  version-specific undo information [4 bytes]
        }

        private struct FibRgFcLcb2000  //FibRgFcLcb2000 structure
        {
            //NOTE: In file begins with FibRgFcLcb97 structure [744 bytes]
            //NOTE: Here is just an additional portion

        }

        private struct FibRgFcLcb2002   //FibRgFcLcb2002 structure
        {
            //NOTE: In file begins with FibRgFcLcb2000 structure [864 bytes]
            //NOTE: Here is just an additional portion

        }

        private struct FibRgFcLcb2003   //FibRgFcLcb2003 structure
        {
            //NOTE: In file begins with FibRgFcLcb2002 structure [1088 bytes]
            //NOTE: Here is just an additional portion

        }

        private struct FibRgFcLcb2007   //FibRgFcLcb2007 structure
        {
            //NOTE: In file begins with FibRgFcLcb2003 structure [1312 bytes]
            //NOTE: Here is just an additional portion

        }
        #endregion

        #region Поля
        #region private

        #endregion

        #region protected internal
        protected internal string FilePath;    //путь к DOC файлу
        #endregion
        #endregion

        #region Свойства
    
        #endregion

        #region Конструкторы
        protected internal docParser(string filePath)
        {
            FilePath = filePath;

            FileStream fileStream = new FileStream(FilePath, FileMode.Open);
            BinaryReader fileReader = new BinaryReader(fileStream, Encoding.Unicode);
                                    
            CompoundFileBinary CFB = new CompoundFileBinary(fileReader);
                        
        }
        #endregion

        #region Методы
        #region private

        #endregion

        #region protected internal

        #endregion
        #endregion
    }
}
