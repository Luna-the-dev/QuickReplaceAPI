/******************************************************************
 * Copyright (c) QuickReplace, LLC. All Rights Reserved.
 * 
 * This file is part of the QuickReplace Developer Library. You may not
 * use this file for commercial or business use except in compliance with
 * the QuickReplace license. You may obtain a copy of the license at:
 * 
 *   https://www.quickreplace.io/pricing
 * 
 * You may view the terms of this license at:
 * 
 *   https://www.quickreplace.io/eula
 * 
 * This file or any others in the QuickReplace Developer Library may not
 * be modified, shared, or distributed outside of the terms of the license
 * agreement by organizations or individuals who are covered by a license.
 * 
 *****************************************************************/

namespace TextReplaceAPI.DataTypes
{
    public record SourceFile
    {
        public string SourceFileName { get; set; }
        public string OutputFileName { get; set; }
        public int NumOfReplacements { get; set; }

        public SourceFile(
            string sourceFileName,
            string outputFileName,
            int numOfReplacements = -1)
        {
            SourceFileName = sourceFileName;
            OutputFileName = outputFileName;
            NumOfReplacements = numOfReplacements;
        }
    }
}
