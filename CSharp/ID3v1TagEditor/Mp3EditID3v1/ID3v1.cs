using System;
using System.IO;
using System.Text;

namespace Mp3EditID3v1
{
	public class ID3v1
	{
		private byte[] _data;

		public ID3v1()
		{
			Bytes = new byte[128];
		}
		public byte[] Bytes
		{
			get{return _data;}
			set{
				_data = new byte[128];
				if(TagExists(value))
					GetBytes(value, value.Length - 128, 128).CopyTo(_data, 0);
				else
					Encoder().GetBytes("TAG").CopyTo(_data, 0);
			}
		}
		protected byte[] GetBytes(byte[] buffer, int start, int count)
		{
			byte[] copy = new byte[count];
			for(int i = 0; i < count; i++) copy[i] = buffer[start + i];
			return copy;
		}
		protected void SetByteString(int start, int count, string value)
		{
			byte[] buffer = Encoder().GetBytes(value);
			for(int i = 0; i < count; i++)
				if(i < buffer.Length)
					_data[start + i] = buffer[i];
				else
					_data[start + i] = 0;
		}
		protected string GetByteString(int index, int length)
		{
			return Encoder().GetString(Bytes, index, length); //.Replace('\0', ' ') .Trim();
		}
		protected void SetByte(int index, byte value)
		{
			_data[index] = value;
		}
		protected byte GetByte(int index)
		{
			return _data[index];
		}
		protected virtual Encoding Encoder()
		{
			return System.Text.Encoding.GetEncoding("iso-8859-1");
		}
		public void Read(string path)
		{
			Stream stream = null;
			try
			{
				stream = File.OpenRead(path);
				Read(stream);
			}
			finally
			{
				if(stream != null) stream.Close();
			}
		}

		public void Read(Stream stream)
		{
			stream.Seek(0, SeekOrigin.Begin);
			byte[] buffer = new byte[stream.Length];
			stream.Read(buffer, 0, buffer.Length);
			Read(buffer);
		}
		
		public void Read(byte[] buffer)
		{
			Bytes = buffer;
		}
		

		public void Write(string path)
		{
			Stream stream = null;
			try
			{
				stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite);
				Write(stream);
			}
			finally
			{
				if(stream != null) stream.Close();
			}
		}
		
		public void Write(Stream stream)
		{
			if(TagExists(stream))
				stream.Seek(-128, SeekOrigin.End);
			else
				stream.Seek(0, SeekOrigin.End);
			stream.Write(Bytes, 0, Bytes.Length);
		}
		
		public void Write(byte[] buffer)
		{
			if(!TagExists(buffer))
			{
				byte[] copy = new byte[buffer.Length + 128];
				buffer.CopyTo(copy, 0);
				buffer = copy;
			}
			Bytes.CopyTo(buffer, buffer.Length - 128);
		}
		
		public bool TagExists(string path)
		{
			Stream stream = null;
			try
			{
				stream = File.OpenRead(path);
				return TagExists(stream);
			}
			finally
			{
				if(stream != null) stream.Close();
			}
		}
		
		public bool TagExists(Stream stream)
		{
			if(stream.Length < 128) return false;
			byte[] signature = new byte[3];
			stream.Seek(-128, SeekOrigin.End);
			stream.Read(signature, 0, 3);
			return Encoder().GetString(signature) == "TAG";
		}
		
		public bool TagExists(byte[] buffer)
		{
			if(buffer.Length < 128) return false;
			return Encoder().GetString(buffer, buffer.Length - 128, 3) == "TAG";
		}
		
		#region TagProperties
		/// <summary>
		/// This field contains the title of the MP3. (30 characters)
		/// </summary>
		public string SongTitle
		{
			get
			{
				return GetByteString(3, 30);
			}
			set
			{
				SetByteString(3, 30, value);
			}
		}
		/// <summary>
		/// This field contains the artist of the MP3 (30 characters).
		/// </summary>
		public string Artist
		{
			get
			{
				return GetByteString(33, 30);
			}
			set
			{
				SetByteString(33, 30, value);
			}
		}
		/// <summary>
		/// This field contains the album where the MP3 comes from (30 characters).
		/// </summary>
		public string Album
		{
			get
			{
				return GetByteString(63, 30);
			}
			set
			{
				SetByteString(63, 30, value);
			}
		}
		/// <summary>
		/// This field contains the year when this song has originally been released. (4 characters)
		/// </summary>
		public string Year
		{
			get
			{
				return GetByteString(93, 4);
			}
			set
			{
				SetByteString(93, 4, value);
			}
		}
		/// <summary>
		/// This field contains a comment for the MP3.  (30 characters - 28 if Track is specified)
		/// </summary>
		public string Comment
		{
			get
			{
				if(_data[125] == 0 && _data[126] != 0)
					return GetByteString(97, 28);
				else
					return GetByteString(97, 30);
			}
			set
			{
				if(_data[125] == 0 && _data[126] != 0)
					SetByteString(97, 28, value);
				else
					SetByteString(97, 30, value);
			}
		}
		/// <summary>
		/// The original album tracknumber
		/// </summary>
		public byte Track
		{
			get
			{
				if(GetByte(125) == 0 && GetByte(126) != 0)
					return GetByte(126);
				else if(Comment.StartsWith("Track "))
					return Convert.ToByte(Comment.Substring(5));
				else
					return 0;
			}
			set
			{
				// TODO: Prevent byte value of 0xFF
				SetByte(125, 0);
				SetByte(126, value);
			}
		}
		
		/// <summary>
		/// This byte contains the offset of a genre in a predefined list.
		/// </summary>
		public byte Genre
		{
			get
			{
				return GetByte(127);
			}
			set
			{
				SetByte(127, value);
			}
		}
		#endregion
		public static string[] Genres()
		{
			return new string[]{
								   "Blues","Classic Rock","Country","Dance","Disco","Funk","Grunge","Hip-Hop","Jazz",
								   "Metal","New Age","Oldies","Other","Pop","R&B","Rap","Reggae","Rock","Techno",
								   "Industrial","Alternative","Ska","Death Metal","Pranks","Soundtrack","Euro-Techno",
								   "Ambient","Trip-Hop","Vocal","Jazz+Funk","Fusion","Trance","Classical","Instrumental",
								   "Acid","House","Game","Sound Clip","Gospel","Noise","AlternRock","Bass","Soul",
								   "Punk","Space","Meditative","Instrumental Pop","Instrumental Rock","Ethnic","Gothic",
								   "Darkwave","Techno-Industrial","Electronic","Pop-Folk","Eurodance","Dream","Southern Rock",
								   "Comedy","Cult","Gangsta","Top 40","Christian Rap","Pop/Funk","Jungle","Native American",
								   "Cabaret","New Wave","Psychadelic","Rave","Showtunes","Trailer","Lo-Fi","Tribal","Acid Punk",
								   "Acid Jazz","Polka","Retro","Musical","Rock & Roll","Hard Rock","Folk","Folk-Rock","National Folk",
								   "Swing","Fast Fusion","Bebob","Latin","Revival","Celtic","Bluegrass","Avantgarde","Gothic Rock",
								   "Progressive Rock","Psychedelic Rock","Symphonic Rock","Slow Rock","Big Band","Chorus",
								   "Easy Listening","Acoustic","Humour","Speech","Chanson","Opera","Chamber Music","Sonata",
								   "Symphony","Booty Bass","Primus","Porn Groove","Satire","Slow Jam","Club","Tango","Samba",
								   "Folklore","Ballad","Power Ballad","Rhythmic Soul","Freestyle","Duet","Punk Rock","Drum Solo",
								   "A capella","Euro-House","Dance Hall"
							   };
		}
	}
}
