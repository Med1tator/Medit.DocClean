using System.IO;
using Aspose.Words;

namespace Medit.DocClean
{
    public class DocCleanUtility
    {
        /// <summary>
        /// doc清洁方法
        /// </summary>
        /// <param name="srcBytes">doc源二进制数据</param>
        /// <param name="extension"></param>
        /// <returns></returns>
        public static byte[] CleanDoc(byte[] srcBytes, string extension)
        {
            // 参数验证
            if (srcBytes == null)
                throw new InvalidDataException();
            if (string.IsNullOrEmpty(extension))
                throw new InvalidDataException();

            byte[] result = null;
            Stream docStream = new System.IO.MemoryStream(srcBytes);
            Aspose.Words.Document doc = new Aspose.Words.Document(docStream);

            // 接受全部修订
            doc.AcceptAllRevisions();

            //StyleCollection styleCollection=

            //清除批注
            NodeCollection nodeCollection = doc.GetChildNodes(NodeType.Comment, true);
            foreach (Node comment in nodeCollection)
            {
                nodeCollection.Remove(comment);
            }

            //二进制另存
            using (MemoryStream ms = new MemoryStream())
            {
                if (extension == ".docx" || extension == "docx")
                {
                    doc.Save(ms, SaveFormat.Docx);
                    result = ms.ToArray();
                }
                if (extension == ".doc" || extension == "doc")
                {
                    doc.Save(ms, SaveFormat.Doc);
                    result = ms.ToArray();
                }
            }
            return result;
        }
    }
}
