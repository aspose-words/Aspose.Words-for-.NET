// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace AsposeVisualStudioPluginWords.Core
{
    public class AsposeComponents
    {
        public static Dictionary<String, AsposeComponent> list = new Dictionary<string, AsposeComponent>();
        public AsposeComponents()
        {
            list.Clear();
            
            AsposeComponent asposeWords = new AsposeComponent();
            asposeWords.set_downloadUrl("");
            asposeWords.set_downloadFileName("aspose.words.zip");
            asposeWords.set_name(Constants.ASPOSE_COMPONENT);
            asposeWords.set_remoteExamplesRepository("https://github.com/asposewords/Aspose_Words_NET.git");
            list.Add(Constants.ASPOSE_COMPONENT, asposeWords);
        }
    }
}
