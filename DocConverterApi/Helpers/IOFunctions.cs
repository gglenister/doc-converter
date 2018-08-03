namespace DocConverterApi.Helpers
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Reflection;

    public static class IOFunctions
    {

        /// <summary>
        ///         ''' Returns a byte array from the stream
        ///         ''' </summary>
        public static byte[] ByteArrayFromStream(System.IO.Stream stream)
        {
            if (stream == null || stream.Length == 0)
                return null;

            // Returns a byte array from the file
            byte[] byteData;
            byteData = new byte[Convert.ToInt32(stream.Length) - 1 + 1];
            stream.Read(byteData, 0, Convert.ToInt32(stream.Length));
            return byteData;
        }

        /// <summary>
        ///         ''' Converts a stream into a string
        ///         ''' </summary>
        public static string StreamToString(System.IO.Stream str)
        {
            if (str != null)
            {
                byte[] buffer = new byte[16385];
                using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
                {
                    int read = str.Read(buffer, 0, buffer.Length);

                    while (read > 0)
                    {
                        ms.Write(buffer, 0, read);
                        read = str.Read(buffer, 0, buffer.Length);
                    }

                    return Convert.ToBase64String(ms.ToArray());
                }
            }

            return null;
        }

        /// <summary>
        ///         ''' Converts a bytea array into a string
        ///         ''' </summary>
        public static string ByteArrayToString(byte[] bytes)
        {
            return Convert.ToBase64String(bytes);
        }

        /// <summary>
        ///         ''' Converts a string into a byte array
        ///         ''' </summary>
        public static byte[] StringToByteArray(string str)
        {
            if (str != null)
                return Convert.FromBase64String(str);

            return null;
        }

        internal static Stream ReadResourceStreamFromExecutingAssembly(string embeddedFileName)
        {
            try
            {
                Assembly executingAssembly = Assembly.GetExecutingAssembly();

                if (executingAssembly != null && !string.IsNullOrEmpty(embeddedFileName))
                {
                    string matchingResourceName = executingAssembly.GetManifestResourceNames().Where(x => !string.IsNullOrEmpty(x)).FirstOrDefault(x => x.IndexOf(embeddedFileName, StringComparison.InvariantCultureIgnoreCase) >= 0);
                    if (matchingResourceName != null)
                        return executingAssembly.GetManifestResourceStream(matchingResourceName);
                }
            }
            catch (Exception ex)
            {
            }

            return null;
        }
    }
}
