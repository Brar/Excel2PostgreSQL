namespace Excel2PostgreSQL
{
    public enum ExitCode
    {
        ErrorUnknown = -1,
        Success = 0,
        ErrorMissingExcelFileArg = 1,
        ErrorFileDoesNotExist = 2,
    }
}
