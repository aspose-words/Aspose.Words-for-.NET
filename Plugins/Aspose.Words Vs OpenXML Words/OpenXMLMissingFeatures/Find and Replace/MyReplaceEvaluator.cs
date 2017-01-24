// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Find_and_Replace
{
    private class MyReplaceEvaluator : IReplacingCallback
    {
        /// <summary>
        /// This is called during a replace operation each time a match is found.
        /// This method appends a number to the match string and returns it as a replacement string.
        /// </summary>
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            e.Replacement = e.Match.ToString() + mMatchNumber.ToString();
            mMatchNumber++;
            return ReplaceAction.Replace;
        }

        private int mMatchNumber;
    }
}
