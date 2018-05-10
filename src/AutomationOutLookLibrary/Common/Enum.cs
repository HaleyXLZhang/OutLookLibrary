namespace AutomationOutLookLibrary
{
    public enum OlItemType
    {
        olMailItem,
        olAppointmentItem,
        olContactItem,
        olTaskItem,
        olJournalItem,
        olNoteItem,
        olPostItem,
        olDistributionListItem,
        olMobileItemSMS = 11,
        olMobileItemMMS = 12
    }

    public enum OlBodyFormat
    {
        olFormatUnspecified,
        olFormatPlain,
        olFormatHTML,
        olFormatRichText
    }

   
}
