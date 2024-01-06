CREATE
    OR

ALTER FUNCTION [dbo].[fn_RTFToText] (@RTF NVARCHAR(MAX))
RETURNS NVARCHAR(MAX)
AS
BEGIN
    /*
     * Given a RTF text, extracts the visible, plain text.
     */
    -- detect RTF text
    IF CHARINDEX(N'{\rtf', @RTF) = 1
    BEGIN
        DECLARE @PlainText NVARCHAR(MAX)
            , @Position INT
            , @Brace NCHAR(1)
            , @Pair INT = 0
            , @PrefixLength INT
            , @i INT
            , @RemainingLength INT
            , @Match NVARCHAR(MAX)
            , @ReplacementChar NVARCHAR(4)
            , @Pattern NVARCHAR(MAX)
            , @TempString NVARCHAR(MAX)
            , @PartialContent NVARCHAR(MAX)
            , @StartPosition INT
            , @EndPosition INT
            , @PairCount INT
            , @Slash NCHAR(2) = CHAR(2) + N'1'
            , @StartBrace NCHAR(2) = CHAR(2) + N'2'
            , @EndBrace NCHAR(2) = CHAR(2) + N'3';
        DECLARE @Segments TABLE (
            Pair INT DEFAULT NULL
            , StartPosition INT DEFAULT NULL
            , EndPosition INT DEFAULT NULL
            , PlainText NVARCHAR(MAX) DEFAULT NULL
            );

        -- remove embedded NULL characters
        SET @RTF = REPLACE(@RTF, CHAR(0), N'');
        -- remove newlines
        SET @RTF = REPLACE(@RTF, CHAR(10), N' ');
        SET @RTF = REPLACE(@RTF, CHAR(13), N' ');
        SET @RTF = REPLACE(@RTF, N'\pard', N'');
        SET @RTF = REPLACE(@RTF, N'\par', N'');
        -- replace escaped characters with temporary symbols: \\ -> CHAR(2)+1; \{ -> CHAR(2)+2; \} -> CHAR(2)+3
        SET @RTF = REPLACE(@RTF, N'\\', @Slash);
        SET @RTF = REPLACE(@RTF, N'\{', @StartBrace);
        SET @RTF = REPLACE(@RTF, N'\}', @EndBrace);
        -- replace the other escaped characters to their literal value: \'[0-9a-f][0-9a-f]  \'FF -> CHAR(1*CONVERT(VARBINARY(MAX), N'0x' + N'FF', 1))
        SET @TempString = @RTF
        SET @Pattern = N'%\''[0-9a-f][0-9a-f]%';
        SET @Position = 1;

        WHILE @Position > 0
        BEGIN
            SET @Position = PATINDEX(@Pattern, @TempString);

            IF @Position > 0
            BEGIN
                SET @Match = SUBSTRING(@TempString, @Position + 2, 2);
                SET @ReplacementChar = CHAR(1 * CONVERT(VARBINARY(MAX), N'0x' + @Match, 1));
                SET @RTF = REPLACE(@RTF, N'\''' + @Match, N'\c ' + @ReplacementChar + N'\');
                SET @TempString = STUFF(@TempString, @Position, 2, N'');
            END;
        END;

        -- replace the Unicode escaped characters to their literal value
        SET @Pattern = N'%\u[0-9]%';
        SET @Position = 1;
        SET @PrefixLength = 2;

        WHILE @Position > 0
        BEGIN
            SET @Position = PATINDEX(@Pattern, @TempString);

            IF @Position > 0
            BEGIN
                SET @RemainingLength = CHARINDEX(N'?', SUBSTRING(@TempString, @Position + @PrefixLength, LEN(@TempString))) - 1;

                IF @RemainingLength > 0
                BEGIN
                    SET @Match = SUBSTRING(@TempString, @Position + @PrefixLength, @RemainingLength);
                    SET @ReplacementChar = NCHAR(@Match);
                    SET @RTF = REPLACE(@RTF, N'\u' + @Match + N'?', N'\c ' + @ReplacementChar + N'\');
                END;

                SET @TempString = STUFF(@TempString, @Position, LEN(@Match) + @PrefixLength + 1, N'');
            END;
        END;

        -- record the positions of {}
        SET @Position = 0;

        WHILE @Position <= LEN(@RTF)
        BEGIN
            SELECT @Position = @Position + 1;

            SELECT @Brace = SUBSTRING(@RTF, @Position, 1);

            IF @Brace = N'{'
            BEGIN
                SELECT @Pair = @Pair + 1;

                INSERT INTO @Segments (
                    Pair
                    , StartPosition
                    , EndPosition
                    )
                VALUES (
                    @Pair
                    , @Position
                    , NULL
                    );
            END;

            IF @Brace = N'}'
            BEGIN
                UPDATE @Segments
                SET EndPosition = @Position
                WHERE Pair = (
                        SELECT MAX(Pair)
                        FROM @Segments
                        WHERE EndPosition IS NULL
                        );
            END;
        END;

        -- remove from output all non-visible inner groups
        SET @TempString = @RTF;

        SELECT @PairCount = COUNT(*)
        FROM @Segments;

        WHILE (@PairCount > 0)
        BEGIN
            SELECT @PairCount = @PairCount - 1;

            SELECT @StartPosition = StartPosition
                , @EndPosition = EndPosition
            FROM @Segments
            ORDER BY Pair DESC OFFSET @PairCount ROWS

            FETCH NEXT 1 ROWS ONLY;

            SET @PartialContent = SUBSTRING(@RTF, @StartPosition, @EndPosition - @StartPosition + 1);

            IF (
                    CHARINDEX(N'{\*', @PartialContent) = 1
                    OR CHARINDEX(N';}', @PartialContent) = (LEN(@PartialContent) - 1)
                    OR CHARINDEX(N' ', @PartialContent) = 0
                    OR CHARINDEX(N'{\info', @PartialContent) = 1
                    )
            BEGIN
                SET @TempString = REPLACE(@TempString, @PartialContent, N'');
            END;
        END;

        -- extract visible texts from the resulting output 
        SET @RTF = @TempString;
        SET @RTF = REPLACE(@RTF, N'{', N'');
        SET @RTF = REPLACE(@RTF, N'}', N'');
        SET @RTF = CONCAT (
                @RTF
                , N'\'
                );
        SET @i = 1;

        -- repeatedly test patterns of type \[^ \] of up to a given amount of characters
        WHILE @i <= 20
        BEGIN
            SET @Pattern = CONCAT (
                    N'%\'
                    , REPLICATE(N'[^ \]', @i)
                    , N' %\%'
                    );
            SET @Position = 1;
            SET @TempString = @RTF;
            SET @StartPosition = 0;

            WHILE @Position > 0
            BEGIN
                SET @Position = PATINDEX(@Pattern, @TempString);

                IF @Position > 0
                BEGIN
                    SET @RemainingLength = CHARINDEX(N'\', SUBSTRING(@TempString, @Position + @i + 2, LEN(@TempString))) - 1;

                    IF @RemainingLength > 0
                    BEGIN
                        SET @Match = SUBSTRING(@TempString, @Position + @i + 2, @RemainingLength);

                        IF LEN(@Match) >= 0
                        BEGIN
                            INSERT INTO @Segments (
                                StartPosition
                                , PlainText
                                )
                            VALUES (
                                @StartPosition + @Position + @i + 2 + LEN(@Match)
                                , @Match
                                );
                        END;
                    END;

                    SET @TempString = SUBSTRING(@TempString, @Position + @i + 2 + LEN(@Match), LEN(@TempString));
                    SET @StartPosition = LEN(@RTF) - LEN(@TempString);
                END;
            END;

            SET @i = @i + 1;
        END;

        SELECT @PlainText = CONCAT (
                @PlainText
                , PlainText
                )
        FROM @Segments
        WHERE Pair IS NULL
        ORDER BY StartPosition;

        -- put back the escaped characters, reduce double spaces and trim the string
        SET @PlainText = REPLACE(@PlainText, @Slash, N'\');
        SET @PlainText = REPLACE(@PlainText, @StartBrace, N'{');
        SET @PlainText = REPLACE(@PlainText, @EndBrace, N'}');

        WHILE CHARINDEX(N'  ', @PlainText) > 0
        BEGIN
            SET @PlainText = REPLACE(@PlainText, N'  ', N' ');
        END;

        SET @PlainText = LTRIM(RTRIM(@PlainText));
    END
    ELSE
    BEGIN
        SET @PlainText = @RTF;
    END;

    RETURN @PlainText;
END;
