using System.Collections.Generic;
using System.Reflection;

namespace MyJournal.Notebook.Config
{
    class AssemblyInfo
    {
        private readonly Assembly _assembly;
        private readonly Dictionary<string, string> _attributes;

        #region Properties

        /// <summary>
        /// Global assembly attributes.
        /// </summary>
        internal IReadOnlyDictionary<string, string> AssemblyAttributes
        { get; private set; }

        /// <summary>
        /// Assembly configuration.
        /// </summary>
        internal string Configuration { get; private set; }

        /// <summary>
        /// Assembly company name.
        /// </summary>
        internal string CompanyName { get; private set; }

        /// <summary>
        /// Assembly copyright.
        /// </summary>
        internal string Copyright { get; private set; }

        /// <summary>
        /// Assembly description.
        /// </summary>
        internal string Description { get; private set; }

        /// <summary>
        /// Assembly product name.
        /// </summary>
        internal string ProductName { get; private set; }

        /// <summary>
        /// Assembly product version.
        /// </summary>
        internal string ProductVersion { get; private set; }

        /// <summary>
        /// Assembly title.
        /// </summary>
        internal string Title { get; private set; }

        /// <summary>
        /// Assembly trademark.
        /// </summary>
        internal string Trademark { get; private set; }

        /// <summary>
        /// Assembly version.
        /// </summary>
        internal string Version { get; private set; }

        #endregion

        internal AssemblyInfo() :
          this(Assembly.GetEntryAssembly() ?? Assembly.GetCallingAssembly())
        { }

        internal AssemblyInfo(Assembly assembly)
        {
            _assembly = assembly;
            _attributes = new Dictionary<string, string>();
            InitializeAssemblyAttributes();
            AssemblyAttributes = _attributes;
        }

        void InitializeAssemblyAttributes()
        {
            SetTitle();
            SetDescription();
            SetConfiguration();
            SetCompanyName();
            SetProductName();
            SetProductVersion();
            SetCopyright();
            SetTrademark();
            SetVersion();
        }

        void SetConfiguration()
        {
            var attribute = _assembly.GetCustomAttribute<AssemblyConfigurationAttribute>();
            Configuration = attribute?.Configuration ?? string.Empty;
            _attributes.Add("Configuration", Configuration);
        }

        void SetCompanyName()
        {
            var attribute = _assembly.GetCustomAttribute<AssemblyCompanyAttribute>();
            CompanyName = attribute?.Company ?? string.Empty;
            _attributes.Add("Company", CompanyName);
        }

        void SetCopyright()
        {
            var attribute = _assembly.GetCustomAttribute<AssemblyCopyrightAttribute>();
            Copyright = attribute?.Copyright ?? string.Empty;
            _attributes.Add("Copyright", Copyright);
        }

        void SetDescription()
        {
            var attribute = _assembly.GetCustomAttribute<AssemblyDescriptionAttribute>();
            Description = attribute?.Description ?? string.Empty;
            _attributes.Add("Description", Description);
        }

        void SetProductName()
        {
            var attribute = _assembly.GetCustomAttribute<AssemblyProductAttribute>();
            ProductName = attribute?.Product ?? string.Empty;
            _attributes.Add("Product", ProductName);
        }

        void SetProductVersion()
        {
            var attribute = _assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
            ProductVersion = attribute?.InformationalVersion ?? string.Empty;
            _attributes.Add("ProductVersion", ProductVersion);
        }

        void SetTitle()
        {
            var attribute = _assembly.GetCustomAttribute<AssemblyTitleAttribute>();
            Title = attribute?.Title ?? string.Empty;
            _attributes.Add("Title", Title);
        }

        void SetTrademark()
        {
            var attribute = _assembly.GetCustomAttribute<AssemblyTrademarkAttribute>();
            Trademark = attribute?.Trademark ?? string.Empty;
            _attributes.Add("Trademark", Trademark);
        }

        void SetVersion()
        {
            Version = _assembly.GetName().Version.ToString();
            _attributes.Add("Version", Version);
        }

        public override string ToString()
        {
            var properties = new string[]
            {
                $"Title           : {Title}",
                $"Description     : {Description}",
                $"Configuration   : {Configuration}",
                $"Company         : {CompanyName}",
                $"Product         : {ProductName}",
                $"Product Version : {ProductVersion}",
                $"Copyright       : {Copyright}",
                $"Trademark       : {Trademark}",
                $"Version         : {Version}"
            };
            return string.Join("\n", properties);
        }
    }
}
