using System.IO;
using System.Text;

namespace EmbeddedText
{
	public class Resources
	{
		private Resources(){}
		public static string GetEmbeddedText(string fileName)
		{
			System.Reflection.Assembly assembly = null;
			Stream stream = null;
			try
			{
				assembly = System.Reflection.Assembly.GetExecutingAssembly();
				fileName = assembly.GetName().Name + "." + fileName;
				stream = assembly.GetManifestResourceStream(fileName);
				byte[] buffer = new byte[stream.Length];
				stream.Read(buffer, 0, buffer.Length);
				return Encoding.ASCII.GetString(buffer);
			}
			finally
			{
				if(stream != null) stream.Close();
			}
		}
	}
}
