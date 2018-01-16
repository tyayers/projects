using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SupportSearch.Common.Interfaces
{
    public interface ISupportSearch
    {
        string SearchMarkdownDocs(string search);
        string SearchHtmlDocs(string search);

    }
}
