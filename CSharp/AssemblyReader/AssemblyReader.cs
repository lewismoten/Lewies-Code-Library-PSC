using System;
using System.IO;
using System.Reflection;

namespace www.LewisMoten.com
{
	/// <summary>Reads assembly information and embedded resources.</summary>
	/// <remarks>Most methods allow you to acquire information from the executing assembly.  However, you may pass a seperate assembly to read the information from.</remarks>
	public class AssemblyReader
	{
		/// <param name="fileName">Then name of the file, including any folders.  If an embedded resource is within a folder in your application, then it is delimited by periods.</param>
		/// <example>The following code demonstrates how you would read an image from the manifest and assign it to a picture box control.
		/// <code lang="C#">Picture.Image = Bitmap.FromStream(AssemblyReader.EmbeddedStream("images.splash.bmp"));</code>
		/// </example>
		/// <summary>Returns a stream of an embedded resource in the executing assembly.</summary>
		static public Stream EmbeddedStream(string fileName)
		{
			return EmbeddedStream(fileName, Assembly.GetExecutingAssembly());
		}
		/// <param name="fileName">Then name of the file, including any folders.  If an embedded resource is within a folder in your application, then it is delimited by periods.</param>
		/// <example>The following code demonstrates how you would read an image from the manifest and assign it to a picture box control.
		/// <code lang="C#">
		/// using System.Reflection;
		/// using System.IO;
		///
		/// Assembly assembly = Assembly.GetExecutingAssembly();
		/// string fileName = "images.splash.bmp";
		/// Stream = AssemblyReader.EmbeddedStream(fileName, assembly);
		/// Picture.Image = Bitmap.FromStream(Stream);
		/// </code>
		/// </example>
		/// <summary>Returns a stream of an embedded resource in the provided assembly.</summary>
		/// <param name="assembly">An assembly containing the embedded file that you wish to acquire.</param>
		static public Stream EmbeddedStream(string fileName, Assembly assembly)
		{
			return assembly.GetManifestResourceStream(assembly.GetName().Name + "." + fileName);
		}
		/// <summary>Returns the AssemblyTitle attribute value commonly defined in the executing assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Title())</code>
		/// </example>
		public static string Title(){return Title(Assembly.GetExecutingAssembly());}
		/// <summary>Returns the AssemblyTitle attribute value commonly defined in the assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Title())</code>
		/// </example>
		/// <param name="assembly">The assembly containing the attribute that you wish to acquire.</param>
		public static string Title(Assembly assembly)
		{
			foreach(object attribute in assembly.GetCustomAttributes(false))
			{
				AssemblyTitleAttribute title = attribute as AssemblyTitleAttribute;
				if(title != null) return title.Title;
			}
			return "";
		}
		/// <summary>Returns the AssemblyDescription attribute value commonly defined in the executing assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Description())</code>
		/// </example>
		public static string Description(){return Description(Assembly.GetExecutingAssembly());}
		/// <summary>Returns the AssemblyDescription attribute value commonly defined in the assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Description())</code>
		/// </example>
		/// <param name="assembly">The assembly containing the attribute that you wish to acquire.</param>
		public static string Description(Assembly assembly)
		{
			foreach(object attribute in assembly.GetCustomAttributes(false))
			{
				AssemblyDescriptionAttribute description = attribute as AssemblyDescriptionAttribute;
				if(description != null) return description.Description;
			}
			return "";
		}
		/// <summary>Returns the AssemblyConfiguration attribute value commonly defined in the executing assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Configuration())</code>
		/// </example>
		public static string Configuration(){return Configuration(Assembly.GetExecutingAssembly());}
		/// <summary>Returns the AssemblyConfiguration attribute value commonly defined in the assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Configuration())</code>
		/// </example>
		/// <param name="assembly">The assembly containing the attribute that you wish to acquire.</param>
		public static string Configuration(Assembly assembly)
		{
			foreach(object attribute in assembly.GetCustomAttributes(false))
			{
				AssemblyConfigurationAttribute configuration = attribute as AssemblyConfigurationAttribute;
				if(configuration != null) return configuration.Configuration;
			}
			return "";
		}
		/// <summary>Returns the AssemblyCompany attribute value commonly defined in the executing assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Company())</code>
		/// </example>
		public static string Company(){return Company(Assembly.GetExecutingAssembly());}
		/// <summary>Returns the AssemblyCompany attribute value commonly defined in the assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Company())</code>
		/// </example>
		/// <param name="assembly">The assembly containing the attribute that you wish to acquire.</param>
		public static string Company(Assembly assembly)
		{
			foreach(object attribute in assembly.GetCustomAttributes(false))
			{
				AssemblyCompanyAttribute company = attribute as AssemblyCompanyAttribute;
				if(company != null) return company.Company;
			}
			return "";
		}
		/// <summary>Returns the AssemblyProduct attribute value commonly defined in the executing assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Product())</code>
		/// </example>
		public static string Product(){return Product(Assembly.GetExecutingAssembly());}
		/// <summary>Returns the AssemblyProduct attribute value commonly defined in the assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Product())</code>
		/// </example>
		/// <param name="assembly">The assembly containing the attribute that you wish to acquire.</param>
		public static string Product(Assembly assembly)
		{
			foreach(object attribute in assembly.GetCustomAttributes(false))
			{
				AssemblyProductAttribute product = attribute as AssemblyProductAttribute;
				if(product != null) return product.Product;
			}
			return "";
		}
		/// <summary>Returns the AssemblyCopyright attribute value commonly defined in the executing assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Copyright())</code>
		/// </example>
		public static string Copyright(){return Copyright(Assembly.GetExecutingAssembly());}
		/// <summary>Returns the AssemblyCopyright attribute value commonly defined in the assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Trademark"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Copyright())</code>
		/// </example>
		/// <param name="assembly">The assembly containing the attribute that you wish to acquire.</param>
		public static string Copyright(Assembly assembly)
		{
			foreach(object attribute in assembly.GetCustomAttributes(false))
			{
				AssemblyCopyrightAttribute copyright = attribute as AssemblyCopyrightAttribute;
				if(copyright != null) return copyright.Copyright;
			}
			return "";
		}
		/// <summary>Returns the AssemblyTrademark attribute value commonly defined in the executing assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Trademark())</code>
		/// </example>
		public static string Trademark(){return Trademark(Assembly.GetExecutingAssembly());}
		/// <summary>Returns the AssemblyTrademark attribute value commonly defined in the assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Version"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Trademark())</code>
		/// </example>
		/// <param name="assembly">The assembly containing the attribute that you wish to acquire.</param>
		public static string Trademark(Assembly assembly)
		{
			foreach(object attribute in assembly.GetCustomAttributes(false))
			{
				AssemblyTrademarkAttribute trademark = attribute as AssemblyTrademarkAttribute;
				if(trademark != null) return trademark.Trademark;
			}
			return "";
		}
		/// <summary>Returns the AssemblyVersion attribute value commonly defined in the executing assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Version())</code>
		/// </example>
		public static string Version(){return Version(Assembly.GetExecutingAssembly());}
		/// <summary>Returns the AssemblyVersion attribute value commonly defined in the assemblies AssemblyInfo file.</summary>
		/// <seealso cref="T:Title"/>
		/// <seealso cref="T:Description"/>
		/// <seealso cref="T:Configuration"/>
		/// <seealso cref="T:Company"/>
		/// <seealso cref="T:Product"/>
		/// <seealso cref="T:Copyright"/>
		/// <seealso cref="T:Trademark"/>
		/// <example>
		/// <code lang="C#">System.Windows.Forms.MessageBox.Show(AssemblyReader.Version())</code>
		/// </example>
		/// <param name="assembly">The assembly containing the attribute that you wish to acquire.</param>
		public static string Version(Assembly assembly)
		{
			foreach(object attribute in assembly.GetCustomAttributes(false))
			{
				AssemblyVersionAttribute version = attribute as AssemblyVersionAttribute;
				if(version != null) return version.Version;
			}
			return "";
		}

	}
}