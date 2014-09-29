using System.IO;

namespace TsSoft.Docx.TemplateEngine
{
    internal static class StreamExtension
    {
        /// <summary>
        /// Since we targeting 3.5, we need custom implementation
        /// </summary>
        /// <param name="input"></param>
        /// <param name="output"></param>
        public static void CopyTo(this Stream input, Stream output)
        {
            byte[] buffer = new byte[16 * 1024];
            int bytesRead;
            while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, bytesRead);
            }
        }
    }
}