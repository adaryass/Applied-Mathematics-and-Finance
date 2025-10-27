# Introduction 
- Une option est un titre financier qui donne à son détenteur le droit et non l'obligation d'acheter ou de vendre un actif sous certains conditions, pendant une période déterminée.

- Il y a plusieurs types d'exercice d'options, les plus connus et les plus utilisés en finance sont principalement trois :
  * **Une option américaine** est une option qui peut être exercée à tout moment jusqu'à la date d'expiration. 
  * **Une option européenne** est une option qui ne peut être exercée qu'à la date de maturité.
  * **Une option bermudéenne** est une option qui ne peut être exercée qu'à certaines dates prédéfinies avant l’échéance.

- Le prix payé pour l'actif lorsque l’option est exercée s’appelle s'appelle **le prix d'exercice** (**striking price** ou **exercice price**). 
- Le dernier jour où l’option peut être exercée s’appelle **la date d’expiration** ou **date d’échéance** (**expiration date** ou **maturity date**).

- Tout d'abord, nous allons aborder **option Call**:
    *  **Boundary Condition**
      
$C_T = [ S_T - K ]^{+}$  = $Max(S_T -K, 0)$


# Black-Scholes Options Pricing Formula

**Modèle Black-Scholes-Merton (BSM)**
  la dynamique du prix de l’action S suivant un _Mouvement Brownien Géométrique_ (GBM), c’est-à-dire un processus lognormal:

$$
S_t = S_0 \exp \left[ \left( \mu - \frac{\sigma^2}{2} \right) t + \sigma W_t \right]
$$

where:
 * µ est le rendement moyen instantané constant (Ici:  **µ = r** taux sans risque; OAT d'un mois).
 * σ  est la volatilité instantanée constante du rendement.
 * $W_t$ est un **processus de Wiener** au temps t, c’est-à-dire **un mouvement brownien**.
   

**Itô's Process** 

D'après Théorème lemme d’Itô:

- Le changement instantané (**Instantaneous price change**) de prix est :
  
$dS_t = \mu S_t dt + \sigma S_t dW_t$

- Le rendement instantané du prix (**Instantaneous price return** est :

$\frac{dS_t}{S_t} = \mu dt + \sigma dW_t$

où :

**$\mu S_t dt$** et **$\sigma S_t dW_t$** sont également couramment appelés **drift term** et **difusion term** associés au processus GBM (Geometric Brownian Motion, ou mouvement brownien géométrique).


**Portfeuille d’arbitrage**

 Considérons un portefeuille d’arbitrage 𝑉, c’est-à-dire un portefeuille répliquant le taux sans risque 𝑟 :

 **𝑉 = C - hS**

 donc:   **d𝑉 = dC - hdS**   avec   **$dC = g(t,S_t)$**

 où :

- 𝑉 est une position couverte sur l’action 𝑆 (position courte) et l’option d’achat 𝐶 (position longue) ;
- 𝑉 ne dépend pas de l’évolution de l’actif sous-jacent : il est delta-couvert (delta-hedged) ;
- $h =\frac{\partial C}{\partial S}$ est le ratio de couverture delta constant dans le temps ;
- Le rendement instantané moyen du portefeuille d’arbitrage est : $\frac{\partial 𝑉}{\ 𝑉} = 𝑟𝑑𝑡$
- La dérive du processus GBM devient 𝜇 = 𝑟 (conséquence de la condition de non-arbitrage).

**Itô's Lemme**
- Itô's Lemma pour dC:
  $dC = \frac{\partial C}{\partial S} dS + \frac{\partial C}{\partial t} dt + \frac{\ 1}{\ 2} \sigma^2 S^2 \frac{\partial^2 C}{\partial S^2} dt$

  $d𝑉 = [ \frac{\partial C}{\partial t} + \frac{\ 1}{\ 2} \sigma^2 S^2 \frac{\partial^2 C}{\partial S^2}] dt$   (a)


  et   $\frac{\partial 𝑉}{\ 𝑉} = 𝑟𝑑𝑡$
  donc :   $d𝑉 = 𝑉𝑟𝑑𝑡 = (C - hS)𝑟𝑑𝑡 = (C - hS)𝑟𝑑𝑡 = (C - \frac{\partial C}{\partial S} S )𝑟𝑑𝑡$  (b)

  (a) = (b) :    **Black-Scholes PDE**
  $𝑟C = 𝑟 \frac{\partial C}{\partial S} S + \frac{\partial C}{\partial t}  + \frac{\ 1}{\ 2} \sigma^2 S^2 \frac{\partial^2 C}{\partial S^2}$

  
  

# Binomial Options Pricing Model of Cox-Ross-Rubinstein
Quantitative finance portfolio: option pricing models (CRR)


```vba
Option Explicit             'Option to declare all the variables: makes your code clearer
                            'You have to declare all your variables

Public Parameters As Variant
Public European_Call() As Double, European_Put() As Double
Public American_Call() As Double, American_Put() As Double
Public S_BinomialTree() As Double
Public S0 As Double, K As Double
Public P As Double, U As Double, D As Double
Public R As Double, T As Double, N As Double
Public i, j As Integer

Public Sub OptionsPricing_CRRModel()

    Dim Wb As Workbook
    Dim Ws1 As Worksheet, Ws2 As Worksheet
    
    Application.ScreenUpdating = False
    
    Set Wb = Workbooks("Lecture1_CRRModel.xlsm")    'Declare the workbook name
    Set Ws1 = Wb.Worksheets("Pricer")               'Declare the first worksheet name
    Set Ws2 = Wb.Worksheets("BinomialTrees")        'Declare the second worksheet name

    'Retrieve parameters
    Application.StatusBar = "Retrieving data..."    'Comment the step in the status bar
   
    With Ws1
        Parameters = Range(.Cells(7, 6), .Cells(7, 6).End(xlDown)).Value    'Input parameters in yellow
    End With
    
    S0 = Parameters(1, 1)       'Stock price at time 0
    K = Parameters(2, 1)        'Strike price
    P = Parameters(3, 1)        'Probability of up (in %)
    U = Parameters(4, 1)        'Delta up per period (in %)
    D = Parameters(5, 1)        'Delta down per period (in %)
    R = Parameters(6, 1)        'Risk-free rate per period (in %)
    T = Parameters(7, 1)        'Time to maturity (in years)
    N = Parameters(8, 1)        '#Periods per year
    
    ReDim European_Call(N * T + 1, N * T + 1) As Double     'Declare the option prices for all the nodes as a matrix
    ReDim European_Put(N * T + 1, N * T + 1) As Double      '#prices: #periods by year * #years + initial time
    ReDim American_Call(N * T + 1, N * T + 1) As Double
    ReDim American_Put(N * T + 1, N * T + 1) As Double
    ReDim S_BinomialTree(N * T + 1, N * T + 1) As Double
    
    'Construct a binomial tree for stock prices
    Application.StatusBar = "Constructing binomial scenariis for stock prices..."
    
    S_BinomialTree(N * T + 1, 1) = S0    'Initial stock price, at time t=0

        For i = 2 To N * T + 1
        'Iteration i is time period of the BT (horizontally): 12 time periods
            S_BinomialTree(N * T + 2 - i, i) = S_BinomialTree(N * T + 3 - i, i - 1) * (1 + U)
            'We compute S only for upward moves: we begin at the bottom of the BT
                For j = N * T + 2 - i To N * T + 1
                'For each time period i, iteration j is the node (vertical prices)
                    If S_BinomialTree(j, i) = 0 Then
                        S_BinomialTree(j, i) = S_BinomialTree(j, i - 1) * (1 - D)
                        'We fill the BT by computing S for downward moves
                    Else
                    End If
                Next j
        Next i
    
    'Calculate European call and put option prices
    Application.StatusBar = "Calculating European call and put option prices..."

        For i = 1 To N * T + 1
            European_Call(i, N * T + 1) = S_BinomialTree(i, N * T + 1) - K
            'For each time period i, we compute the call payoff as stock price minus strike price
                If European_Call(i, N * T + 1) < 0 Then
                'If the call is OTM, ...
                    European_Call(i, N * T + 1) = 0
                    '... then its value is zero: no exercise
                Else
                End If
        Next i
        
        For j = 1 To N * T + 1
        'For each time period j, ...
            For i = j + 1 To N * T + 1
            'For each node i of each time period j, ...
                European_Call(i, N * T + 1 - j) = _
                    (P * European_Call(i - 1, N * T + 2 - j) + _
                    (1 - P) * European_Call(i, N * T + 2 - j)) / (1 + R)
                '... We compute call payoff as in Equation (1) following arbitrage portfolio
            Next i
        Next j
    
    Ws1.Cells(7, 9).Value = European_Call(N * T + 1, 1)
    'The pricer gives the initial call price at time t=0

        For i = 1 To N * T + 1
            European_Put(i, N * T + 1) = K - S_BinomialTree(i, N * T + 1)
            'For each time period i, we compute the put payoff as stock price minus strike price
                If European_Put(i, N * T + 1) < 0 Then
                'If the put is OTM, ...
                    European_Put(i, N * T + 1) = 0
                    '... then its value is zero: no exercise
                Else: End If
        Next i
        
        For j = 1 To N * T + 1
        'For each node j, ...
            For i = j + 1 To N * T + 1
            'For each time period i, ...
                European_Put(i, N * T + 1 - j) = _
                    (P * European_Put(i - 1, N * T + 2 - j) + _
                    (1 - P) * European_Put(i, N * T + 2 - j)) / (1 + R)
                    '... We compute put payoff as in Equation (1) following arbitrage portfolio
            Next i
        Next j
    
    Ws1.Cells(8, 9).Value = European_Put(N * T + 1, 1)
    'The pricer gives the initial put price at time t=0
    
    'Calculate American call and put option prices
    Application.StatusBar = "Calculating American call and put option prices..."
    
        For i = 1 To N * T + 1
            American_Call(i, N * T + 1) = S_BinomialTree(i, N * T + 1) - K
            'For each time period i, we compute the put payoff as stock price minus strike price
                If American_Call(i, N * T + 1) < 0 Then
                'If the call is OTM, ...
                    American_Call(i, N * T + 1) = 0
                    '... then its value is zero: no exercise
                Else
                End If
        Next i
        
        For j = 1 To N * T + 1
        'For each node j, ...
            For i = j + 1 To N * T + 1
            'For each time period i, ...
                American_Call(i, N * T + 1 - j) = _
                    (P * American_Call(i - 1, N * T + 2 - j) + _
                    (1 - P) * American_Call(i, N * T + 2 - j)) / (1 + R)
                    '... We compute call payoff as in Equation (1) following arbitrage portfolio
                        If S_BinomialTree(i, N * T + 1 - j) - K > American_Call(i, N * T + 1 - j) Then
                            American_Call(i, N * T + 1 - j) = S_BinomialTree(i, N * T + 1 - j) - K
                            'Since American options can be exercised at any time, they are worth AT LEAST
                            'as much as European options
                        Else
                        End If
            Next i
        Next j
    
    Ws1.Cells(9, 9).Value = American_Call(N * T + 1, 1)
    'The pricer gives the initial call price at time t=0
    
        For i = 1 To N * T + 1
            American_Put(i, N * T + 1) = K - S_BinomialTree(i, N * T + 1)
            'For each time period i, we compute the put payoff as stock price minus strike price
                If American_Put(i, N * T + 1) < 0 Then
                'If the put is OTM, ...
                    American_Put(i, N * T + 1) = 0
                    '... then its value is zero: no exercise
                Else
                End If
        Next i
        
        For j = 1 To N * T + 1
        'For each node j, ...
            For i = j + 1 To N * T + 1
            'For each time period i, ...
                American_Put(i, N * T + 1 - j) = _
                    (P * American_Put(i - 1, N * T + 2 - j) + _
                    (1 - P) * American_Put(i, N * T + 2 - j)) / (1 + R)
                    '... We compute put payoff as in Equation (1) following arbitrage portfolio
                        If K - S_BinomialTree(i, N * T + 1 - j) > American_Put(i, N * T + 1 - j) Then
                            American_Put(i, N * T + 1 - j) = K - S_BinomialTree(i, N * T + 1 - j)
                            'Since American options can be exercised at any time, they are worth AT LEAST
                            'as much as European options
                        Else
                        End If
            Next i
        Next j
    
    Ws1.Cells(10, 9).Value = American_Put(N * T + 1, 1)
    'The pricer gives the initial put price at time t=0
    
    'Print results
    Application.StatusBar = "Printing results..."   'We print BT prices
    
    With Ws2
    
        .Range("A:XEZ").ClearContents
        .Activate
        .Cells(1, 1).Select
    
        ActiveCell.Range("A1").Value = "Stock Price"
'        With ActiveCell.Range(Cells(1, 1), Cells(1, N * T + 1))
'            .Merge
'            .HorizontalAlignment = xlCenter
'            .Font.Size = "14"
'            .Font.Bold = True
'        End With
        For i = 1 To N * T + 1
            ActiveCell.Cells(2, i).Value = "Time " & (i - 1)
        Next i
        For i = 1 To N * T + 1
            For j = 1 To N * T + 1
                ActiveCell.Cells(i + 2, j) = S_BinomialTree(i, j)
            Next j
        Next i
        ActiveCell.Offset(N * T + 4, 0).Select
    
        ActiveCell.Range("A1").Value = "European Call Price"
'        With ActiveCell.Range(Cells(1, 1), Cells(1, N * T + 1))
'            .Merge
'            .HorizontalAlignment = xlCenter
'            .Font.Size = "14"
'            .Font.Bold = True
'        End With
        For i = 1 To N * T + 1
            ActiveCell.Cells(2, i).Value = "Time " & (i - 1)
        Next i
        For i = 1 To N * T + 1
            For j = 1 To N * T + 1
                ActiveCell.Cells(i + 2, j) = European_Call(i, j)
            Next j
        Next i
        ActiveCell.Offset(N * T + 4, 0).Select
    
        ActiveCell.Range("A1").Value = "European Put Price"
'        With ActiveCell.Range(Cells(1, 1), Cells(1, N * T + 1))
'            .Merge
'            .HorizontalAlignment = xlCenter
'            .Font.Size = "14"
'            .Font.Bold = True
'        End With
        For i = 1 To N * T + 1
            ActiveCell.Cells(2, i).Value = "Time " & (i - 1)
        Next i
        For i = 1 To N * T + 1
            For j = 1 To N * T + 1
                ActiveCell.Cells(i + 2, j) = European_Put(i, j)
            Next j
        Next i
        ActiveCell.Offset(N * T + 4, 0).Select
        
        ActiveCell.Range("A1").Value = "American Call Price"
'        With ActiveCell.Range(Cells(1, 1), Cells(1, N * T + 1))
'            .Merge
'            .HorizontalAlignment = xlCenter
'            .Font.Size = "14"
'            .Font.Bold = True
'        End With
        For i = 1 To N * T + 1
            ActiveCell.Cells(2, i).Value = "Time " & (i - 1)
        Next i
        For i = 1 To N * T + 1
            For j = 1 To N * T + 1
                ActiveCell.Cells(i + 2, j) = American_Call(i, j)
            Next j
        Next i
        ActiveCell.Offset(N * T + 4, 0).Select
        
        ActiveCell.Range("A1").Value = "American Put Price"
'        With ActiveCell.Range(Cells(1, 1), Cells(1, N * T + 1))
'            .Merge
'            .HorizontalAlignment = xlCenter
'            .Font.Size = "14"
'            .Font.Bold = True
'        End With
        For i = 1 To N * T + 1
            ActiveCell.Cells(2, i).Value = "Time " & (i - 1)
        Next i
        For i = 1 To N * T + 1
            For j = 1 To N * T + 1
                ActiveCell.Cells(i + 2, j) = American_Put(i, j)
            Next j
        Next i
        ActiveCell.Offset(N * T + 4, 0).Select
        
    End With

    With Ws1
        .Activate
        Range(Cells(7, 9), .Cells(7, 9).End(xlDown)).NumberFormat = "0.00"
    End With
    
    Application.StatusBar = ""
    Application.ScreenUpdating = True

End Sub

```
