using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace AdrkhaTypograph
{
    public static class OpenTypeParser
    {
        public static List<string> GetFeatures(string fontPath)
        {
            var features = new List<string>();
            if (string.IsNullOrEmpty(fontPath) || !File.Exists(fontPath))
                return features;

            try
            {
                using (var fs = new FileStream(fontPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var br = new BinaryReader(fs))
                {
                    // 1. Read Offset Table (12 bytes)
                    byte[] sfntVersion = br.ReadBytes(4);
                    ushort numTables = ReadUInt16BE(br);
                    fs.Position += 6; // Skip searchRange (2), entrySelector (2), rangeShift (2)

                    // 2. Find GSUB table record
                    uint gsubOffset = 0;
                    uint gsubLength = 0;
                    for (int i = 0; i < numTables; i++)
                    {
                        byte[] tagBytes = br.ReadBytes(4);
                        string tag = Encoding.ASCII.GetString(tagBytes);
                        uint checkSum = ReadUInt32BE(br);
                        uint offset = ReadUInt32BE(br);
                        uint length = ReadUInt32BE(br);

                        if (tag == "GSUB")
                        {
                            gsubOffset = offset;
                            gsubLength = length;
                            break;
                        }
                    }

                    if (gsubOffset == 0 || gsubLength == 0)
                        return features;

                    // 3. Read GSUB Header
                    fs.Position = gsubOffset;
                    ushort majorVersion = ReadUInt16BE(br);
                    ushort minorVersion = ReadUInt16BE(br);
                    ushort scriptListOffset = ReadUInt16BE(br);
                    ushort featureListOffset = ReadUInt16BE(br);
                    ushort lookupListOffset = ReadUInt16BE(br);

                    // 4. Read FeatureList Table
                    // FeatureList table is located at GSUB start + featureListOffset
                    fs.Position = gsubOffset + featureListOffset;
                    ushort featureCount = ReadUInt16BE(br);

                    // Read FeatureRecords (4-byte tag + 2-byte offset)
                    for (int i = 0; i < featureCount; i++)
                    {
                        byte[] tagBytes = br.ReadBytes(4);
                        string tag = Encoding.ASCII.GetString(tagBytes).Trim();
                        ushort featureOffset = ReadUInt16BE(br);

                        if (!string.IsNullOrEmpty(tag) && !features.Contains(tag))
                        {
                            features.Add(tag);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // In case of any binary parsing error, fail silently and return empty features list.
                System.Diagnostics.Debug.WriteLine("Error parsing OpenType features: " + ex.Message);
            }

            return features;
        }

        private static ushort ReadUInt16BE(BinaryReader br)
        {
            byte[] bytes = br.ReadBytes(2);
            if (bytes.Length < 2) return 0;
            if (BitConverter.IsLittleEndian)
                Array.Reverse(bytes);
            return BitConverter.ToUInt16(bytes, 0);
        }

        private static uint ReadUInt32BE(BinaryReader br)
        {
            byte[] bytes = br.ReadBytes(4);
            if (bytes.Length < 4) return 0;
            if (BitConverter.IsLittleEndian)
                Array.Reverse(bytes);
            return BitConverter.ToUInt32(bytes, 0);
        }
    }
}
