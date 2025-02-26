Attribute VB_Name = "Module1"
Option Explicit
Option Base 0

'Declare all variables that will be used across submodules
Dim totalincome As Currency
Dim homeprice As Currency
Dim proptaxes As Currency
Dim monthpayinput As Currency
Dim downpaymentinput As Currency
Dim percentdown As Double
Dim years As Integer
Dim interest As Double
Dim effectiverate As Double
Dim heat As Currency
Dim condo As Currency
Function FindInsuranceValue(percentdown, downparray As Range, insurancearray As Range, mortgage, AmortPeriod)

'Find Insurance Value
FindInsuranceValue = (WorksheetFunction.XLookup(percentdown, downparray, insurancearray, 0, -1) * mortgage) / AmortPeriod

End Function
Function AnnualPayment(effective_rate As Double, years As Integer, mortgage)

'Find Annual Payment Amount
AnnualPayment = (WorksheetFunction.Pmt(effective_rate, years * 12, mortgage) * -12)

End Function
Sub InitialUserInput()

Dim income As Currency
Dim debt As Currency

'Input and place the user's income
income = InputBox("Please enter your annual income." & Chr(10) & "Example: 97000", "Income - Maximum Mortgage Calculator")
Range("AnnualIncome").Value = income

'Input and place the user's debt
debt = InputBox("Please enter your monthly debt excluding credit cards." & Chr(10) & "Example: 1000", "Debt - Maximum Mortgage Calculator")
Range("MonthlyDebt").Value = debt

'Calculate total income
totalincome = income - (12 * debt)
Range("TotalIncome").Value = totalincome

'Input and place how much the user is willing to pay per month
monthpayinput = InputBox("How much are you willing to pay in total house costs (mortgage, heating, insurance etc) per month?" & Chr(10) & "Example: 2000", "Maximum Monthly Pay - Maximum Mortgage Calculator")
Range("maxmonthly") = monthpayinput

'Input and place the cost of the home
homeprice = InputBox("How much is the home you're looking to buy?" & Chr(10) & "Example: 500000", "Maximum Mortgage Calculator")
Range("homeprice").Value = homeprice

'Calculate and place property taxes
proptaxes = homeprice * 0.0065718
Range("PropTaxes").Value = proptaxes

End Sub

Sub DownPayment()

Dim mindown As Currency

'If home price is less than 500,000 then find the minimum downpayment of 5% of the homeprice
If homeprice < 500000 Then
    mindown = homeprice * 0.05
'If the homeprice is less than 1,000,000 the find the minimum downpayment of 5% of 500000 and 10% of the remaining amount
ElseIf homeprice < 1000000 Then
    mindown = (500000 * 0.05) + (homeprice - 500000) * 0.1
'If the homeprice is over a $1,000,000 find the minimum downpayment of 20% of the homeprice
Else
    mindown = homeprice * 0.2
End If


'Insert Downpayment
downpaymentinput = InputBox("How much are you willing to pay as a downpayment? Please enter the monetary amount." & Chr(10) & "Ex $10000." & Chr(10) & "The minimum you must pay is " & Format(mindown, "Currency"), "Down Payment - Maximum Mortgage Calculator")

'Ensure the inputted downpayment is at least the minimum downpayment, otherwise inform the user of the minimum downpayment and repeat the question until it reaches the minimum
If downpaymentinput < mindown Then
    Do Until downpaymentinput >= mindown
        downpaymentinput = InputBox("You require a minimum downpayment of " & Format(mindown, "Currency") & " to purchase this house, please try again")
    Loop
    'Place corrected downpayment if necessary
    Range("downpayment").Value = downpaymentinput
'Place the userinputted downpayment
Else: Range("downpayment").Value = downpaymentinput

End If

'Calculate and place the percentage of the home used as a downpayment
percentdown = downpaymentinput / homeprice
Range("percentdown").Value = percentdown

End Sub

Sub AmortPeriod()

'Create maximum amortization period variable
Dim maxamort As Integer

'identify maximum amortization period, depending on the downpayment percentage put down on the home so the user can be informed
If percentdown < 0.2 Then
    maxamort = 25
Else
    maxamort = 35
End If


' Store the User's for initial amortization period Input with the maximum amortization period warning provided
years = InputBox("Please enter the Amortization Period (Years). Please note that for this home, the maximum amortization period possible is " & maxamort & " years", "Amortization Period - Maximum Mortgage Calculator")

'If the home must have insurance, and the user inputted a number greater than 25, loop until the number is less than 25
If percentdown < 0.2 And years > 25 Then
    Do Until years <= 25
        years = InputBox("Uninsured mortgages can only be amortized for up to 25 years")
    Loop
    'Place is corrected amortization period
    Range("AmortPeriod").Value = years

'Place the amortization period if the user did not need correction
Else: Range("AmortPeriod").Value = years

End If

'If the user enters a number greater than 35, loop until corrected
If years > 35 Then
    Do Until years <= 35
        years = InputBox("Mortgages can only be amortized for up to 35 years")
    Loop
    Range("AmortPeriod").Value = years
    
'Place the amortization period if the user did not need correction
Else: Range("AmortPeriod").Value = years

End If

End Sub

Sub InterestRate()

'Set Interest Variables
Dim interestbenchmark As Double
Dim interestQ As VbMsgBoxResult
Dim interestQ2 As VbMsgBoxResult


'Set Interest Benchmark
interestbenchmark = 0.0525

' Ask user if they want to use benchmark rate or their own
interestQ = MsgBox("Would you like to use the benchmark rate of 5.25%?", vbYesNo + vbQuestion + vbDefaultButton1, "Interest Rate - Maximum Mortgage Calculator")

'If the user chooses to use the benchmark rate, place the benchmark rate and recognize it as the interest value
If interestQ = vbYes Then
    interest = interestbenchmark
    Range("InterestRate").Value = interest
'If the user chooses to insert their own rate, ask them to confirm their decision
ElseIf interestQ = vbNo Then
    interestQ2 = MsgBox("Would you like to enter a contractual rate you have recieved? Please note that the greater number between the contractual rate provided plus 2 percentage points and 5.25% will be used", vbYesNo + vbQuestion + vbDefaultButton1, "Interest Rate - Maximum Mortgage Calculator")
    'If they still choose to insert their own rate, increase the amount provided by 2%
    If interestQ2 = vbYes Then
        interest = InputBox("Please enter your contractual rate as a decimal number" & Chr(10) & "Example: 0.054") + 0.02
        'If the amount provided is greater than the benchmark, inform the user and set it as the interest rate
        If interest > interestbenchmark Then
          MsgBox "The provided rate provided plus the additional 2 percentage points has led to the rate of " & Format(interest, "Percent") & ". This interest rate will be used", vbExclamation
          Range("InterestRate").Value = interest
          
        'if the amount is less than the benchmark, inform the user the benchmark will be used and set the benchmark as the interest
        Else
            MsgBox "The benchmark rate of " & Format(interestbenchmark, "Percent") & " will be used", vbExclamation
            interest = interestbenchmark
            Range("InterestRate").Value = interest
        End If
    'if they change their mind, inform the user, use the benchmark and set the benchmark as the interest
    ElseIf interestQ2 = vbNo Then
        interest = interestbenchmark
        Range("InterestRate").Value = interest
        MsgBox "The benchmark rate of " & Format(interestbenchmark, "Percent") & " will be used", vbExclamation
    End If
End If

'calculate and place effective rate
effectiverate = (1 + interest / 2) ^ (1 / 6) - 1
Range("Effective_Rate").Value = effectiverate

End Sub

Sub HeatCost()

'Store the User's Input
heat = InputBox("Please enter your estimated annual heating costs" & Chr(10) & "Example: 1800", "Heating Costs - Maximum Mortgage Calculator")

' Insert Heating Costs
Range("HeatCosts").Value = heat

End Sub

Sub CondoFees()

'Set msgbox variable
Dim congo_msg As VbMsgBoxResult

' Store the User's Response
congo_msg = MsgBox("Are you planning on purchasing a condo?", vbYesNo + vbQuestion + vbDefaultButton2, "Condo Fees - Maximum Mortgage Calculator")

'If the user has condo fees, calculate 50% of that an insert, otherwise insert 0
If congo_msg = vbYes Then
    condo = InputBox("Please enter your estimated annual condo fees", "Maximum Mortgage Calculator") * 0.5
    Range("condofees").Value = condo
Else
    condo = 0
    Range("CondoFees").Value = condo
    
End If
End Sub
Sub MaxMortgage()

'set variables
Dim PITH As Double
Dim insurance As Currency
Dim annualmortgage As Currency
Dim mortgage As Currency
Dim totalmonthlycost As Currency
Dim maxhomeprice As Currency

'sets the dynamic value of insurance and annual mortgage using created functions so they can be changed during goal seek
Range("insurance").Value = "=FindInsuranceValue(percentdown,downparray,insurancearray,mortgage,amortperiod)"
Range("annualmortgage").Value = "=AnnualPayment(effective_rate,amortperiod,mortgage)"

'sets a formula in PITH cell so it can be used during goal seek
Range("PITH").Value = "=(SUM(annualmortgage,proptaxes,insurance,CondoFees,HeatCosts))/TotalIncome"

'use goal seek to find max mortgage
Range("PITH").GoalSeek Goal:=0.32, ChangingCell:=Range("mortgage")

'save max mortgage value
mortgage = Range("mortgage").Value

'remove formula from PITH cell by copying and paste as values
Range("PITH").Copy
Range("PITH").PasteSpecial xlPasteValues
PITH = Range("PITH").Value

'remove formula from Insurance cell by copying and paste as values
Range("insurance").Copy
Range("insurance").PasteSpecial xlPasteValues
insurance = Range("insurance").Value

'remove formula from Insurance cell by copying and paste as values
Range("annualmortgage").Copy
Range("annualmortgage").PasteSpecial xlPasteValues
annualmortgage = Range("annualmortgage").Value

'calculate total monthly cost needed to afford house and place in sheet
totalmonthlycost = (proptaxes + heat + condo + insurance + annualmortgage) / 12
Range("monthlyneeded").Value = totalmonthlycost

'calculate the maximum home price the user can afford
maxhomeprice = mortgage + downpaymentinput
Range("maxhomeprice").Value = maxhomeprice

'if the maximum home price is less than the homeprice inputted, the user cannot afford the house and will be informed of what they can afford instead
If maxhomeprice < homeprice Then
    MsgBox "You cannot afford this home." & Chr(10) & "You can only afford taking out a mortgage of " & Format(mortgage, "Currency") & Chr(10) & " The maximum home price you can afford is " & Format(maxhomeprice, "Currency"), vbExclamation, "Result - Maximum Mortgage Calculator"
'if the total monthly cost needed to afford the home is greater than the monthly cost the user was willing to pay, they will be informed that they can afford the home but must pay more
ElseIf totalmonthlycost > monthpayinput Then
    MsgBox "Good news! You can afford this house!" & Chr(10) & "Bad news! You must be willing to spend " & Format(totalmonthlycost - monthpayinput, "Currency") & " more per month than you planned" & Chr(10) & "You will be approved for a mortgage of up to " & Format(mortgage, "currency") & Chr(10) & "You can afford any home of up to " & Format(maxhomeprice, "Currency"), vbInformation, "Result - Maximum Mortgage Calculator"
'if the maximum home price is greater than the homeprice inputted, the user can afford the home and is informed
ElseIf maxhomeprice >= homeprice Then
    MsgBox "Congratulations! You can afford this house!" & Chr(10) & "Your monthly cost will be " & Format(totalmonthlycost, "currency") & Chr(10) & "You will be approved for a mortgage of up to " & Format(mortgage, "currency") & Chr(10) & "You can afford any home of up to " & Format(maxhomeprice, "Currency"), vbInformation, "Result - Maximum Mortgage Calculator"

End If


End Sub

Sub RunCalculator()

'Call Sub Modules
Call InitialUserInput
Call DownPayment
Call AmortPeriod
Call InterestRate
Call HeatCost
Call CondoFees
Call MaxMortgage

End Sub

Sub ClearCalc()

'Clear all filled rows of the calculator
Range("6:6").ClearContents
Range("9:9").ClearContents
Range("12:12").ClearContents
Range("15:15").ClearContents
Range("18:18").ClearContents
Range("21:21").ClearContents
Range("24:24").ClearContents
Range("27:27").ClearContents
Range("30:30").ClearContents
Range("33:33").ClearContents

'Confirm when clearing completed
MsgBox "The Calculator has been reset :)", vbInformation, "Clear Calculator"

End Sub

