table 50202 "Closing Stock Details Buffer"
{
    Caption = 'Closing Stock Details';
    fields
    {
        field(1; "Entry No."; Integer)
        {
            Caption = 'Entry No.';
        }
        field(2; "Item No."; Text[250])
        {
            Caption = 'Item No.';
        }
        field(3; "Location Code"; Text[250])
        {
            Caption = 'Location Code';
        }
        field(4; "Entry Date"; Date)
        {
            Caption = 'Entry Date';
        }
        field(5; "Stock Quantity"; Decimal)
        {
            Caption = 'Stock Quantity';
        }
        field(6; "Net Quantity"; Decimal)
        {
            Caption = 'Net Quantity';
        }
        field(7; "Closing Detail status"; Option)
        {
            Caption = 'Closing status';
            OptionMembers = " ",Processed,Error,"On Hold";
        }
        field(8; "Previous Quantity"; Decimal)
        {
            Caption = 'Previuos Quantity';
        }
        field(9; "Bin Code"; Text[250])
        {
            Caption = 'Bin Code';
        }
        field(10; "Excel Line No."; Integer)
        {
            Caption = 'Excel Line No.';
        }
    }
    keys
    {
        key(PK; "Entry Date", "Item No.", "Location Code", "Bin Code")
        {
        }
    }
}