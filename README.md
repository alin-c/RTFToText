**fn_RTFToText** is a T-SQL function which extracts from a RTF input the plain text visible to the user.
It handles escaped characters (ASCII, Unicode and \{, \}, \\) so that the output displays them correctly.
It does not handle image data, nor large inputs (larger than NVARCHAR(MAX)).
