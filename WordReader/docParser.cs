using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            #endregion

            #region Поля
            #region private
            private CompoundFileHeader CFHeader;    //Compound file header
            private uint[] DIFAT;                   //entire DIFAT array (from header + from DIFAT sectors)
            private uint[] FAT;                     //FAT array
            private uint[] miniFAT;                 //miniFAT array (from standart chain from header and FAT)
            #endregion

            #region protected internal
            BinaryReader fileReader;    //doc binary reader
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

            private void showBytesInHEX(byte[] Output,string title="",string ending="") //вывод байтового массива в виде HEX
                                                                                        //title - заголовок перед выводом, ending - строка после вывода
            {
                int row = 0;
                int byteNumber = 0;
                Console.WriteLine("\t" + title);
                Console.Write($"{byteNumber:X6}: ");
                foreach (byte o in Output)
                {
                    if (byteNumber != 0 && (byteNumber % 16) == 0)  //вывели 16 байт
                    {
                        Console.WriteLine();
                        row++;
                        Console.Write($"{byteNumber:X6}: ");
                    }
                    Console.Write($"{o:X2}");
                    byteNumber++;
                    if (byteNumber % 2 == 0) Console.Write(" ");
                }

                Console.WriteLine("\n\t" + ending);
            }

            private bool readCFHeader() //читает Compound file header из fileReader и проверяет его на адекватность
                //вернет true, если заголовок без ошибок
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

                //проверяем на ошибки
                //эталонные (MUST) значения полей заголовка
                byte[] signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
                byte[] minorVersion = { 0x3E, 0x00 };
                byte[][] majorVersion = { new byte[] { 0x03, 0x00 }, new byte[] { 0x04, 0x00 } };
                byte[] byteOrder = { 0xFE, 0xFF };
                byte[][] sectorShift = { new byte[] { 0x09, 0x00 }, new byte[] { 0x0C, 0x00 } };
                byte[] miniSectorShift = { 0x06, 0x00 };

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

                //пропускаем остаток сектора, если Major Version is 4
                if (CFHeader.MajorVersion.SequenceEqual(majorVersion[1])) fileReader.BaseStream.Seek(3584, SeekOrigin.Current);

                return true;
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

                miniFAT = new uint[CFHeader.NumMiniFATSectors]; //выделили память под miniFAT
                uint currentminiFATsector=CFHeader.FirstMiniFATSectorLoc;  //номер текущего сектора miniFAT

                for (int i = 0; i < CFHeader.NumMiniFATSectors; i++)
                {
                    miniFAT[i] = currentminiFATsector;  //сохранили номер текущего miniFAT сектора
                    currentminiFATsector = FAT[currentminiFATsector];   //берем из FAT номер следующего miniFAT сектора
                }
            }
            #endregion

            #region protected internal

            #endregion
            #endregion
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
