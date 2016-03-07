' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' Set the appropriate mail merge clean up options to remove any unused regions from the document.
doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions
' doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields
' doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions Or MailMergeCleanupOptions.RemoveStaticFields
' doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions Or MailMergeCleanupOptions.RemoveEmptyParagraphs
' doc.MailMerge.CleanupOptions = doc.MailMerge.CleanupOptions Or MailMergeCleanupOptions.RemoveUnusedFields
