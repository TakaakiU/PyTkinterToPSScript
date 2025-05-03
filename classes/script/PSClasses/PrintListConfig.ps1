# Custom class definition
class PrintListConfig {
    [System.String]$Path
    [System.Object[]]$HeaderConstants
    [System.Object[]]$HeaderValues
    [System.Object[]]$MainConstants
    [System.Object[]]$MainValues

    # Constants (managed within the class)
    static [System.String]$TEMPLATESHEET1 = '1st Page'
    static [System.Int32]$TEMPLATEROWS1 = 30
    static [System.String]$TEMPLATESHEET2 = '2nd Page and beyond'
    static [System.Int32]$TEMPLATEROWS2 = 35
    static [System.String]$HEADERRANGE = 'B2:AA9'
    static [System.String]$MAINRANGE = 'B7:AA41'

    PrintListConfig([System.String]$Path, [System.Object[]]$HeaderConstants, [System.Object[]]$HeaderValues, [System.Object[]]$MainConstants, [System.Object[]]$MainValues) {
        $this.Path = $Path
        $this.HeaderConstants = $HeaderConstants
        $this.HeaderValues = $HeaderValues
        $this.MainConstants = $MainConstants
        $this.MainValues = $MainValues
    }
}
