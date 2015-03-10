namespace TsSoft.Docx.TemplateEngine
{
    using TsSoft.Docx.TemplateEngine.Tags;

    public enum SdtTagLockingType
    {
        [LockValue("unlocked")]
        Unlocked = 0,

        [LockValue("sdtLocked")]
        SdtLocked = 1,

        [LockValue("contentLocked")]
        ContenLocked = 2,

        [LockValue("sdtContentLocked")]
        StdtContentLocked = 3,
    }
}
