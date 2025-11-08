<#
PS4DS: Optimize
Author: Eric K. Miller
Last updated: 7 November 2025

This script contains PowerShell code for optimizing (modeling) data.
#>

#========================
#   Optimize functions
#========================

using namespace MathNet.Numerics.LinearAlgebra

# Since we downloaded from the internet, we will need to unblock
# the file before we can use it (may need administrator permissions).

#Unblock-File -Path $MathNetNumerics_dll

Add-Type -Path $MathNetNumerics_dll

function Measure-CostFunction {
    <#
    .SYNOPSIS
        Calculate the value of the cost function for linear regression.

    .DESCRIPTION
        Given X matrix, Y vector, and Theta parameters, calculate the
    value of the cost function to be used in gradient descent.
    
    .PARAMETER X
        A matrix (of dimension (mx2)) of the independent variable
    (the 2 is for accounting for the bias/intercept term).
    
    .PARAMETER Y
        A vector (of dimension (mx1)) of the dependent variable.
    
    .PARAMETER Theta
        A vector (of dimension (2x1)) of the slope and intercept
    parameters.
    
    .EXAMPLE
        $cost = Measure-CostFunction -X $X -Y $Y -Theta $Theta
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]$X,
        [Parameter(Mandatory)]$Y,
        [Parameter(Mandatory)]$Theta
    )
    
    $m = $Y.Count
    $predictions = $X * $Theta
    $cost = (1/(2*$m)) * ($predictions - $Y).PointwisePower(2).Sum()
    return [double]$cost
}

function Use-GradientDescent {
    <#
    .SYNOPSIS
        Run gradient descent to calculate parameters for linear
    regression.

    .DESCRIPTION
        Given X matrix, Y vector, Theta parameters, LearningRate, and
    Iterations, run gradient descent to calculate and update parameters
    and cost values to determine the solution to a linear regression
    model.
    
    .PARAMETER X
        A matrix (of dimension (mx2)) of the independent variable
    (the 2 is for accounting for the bias/intercept term).
    
    .PARAMETER Y
        A vector (of dimension (mx1)) of the dependent variable.
    
    .PARAMETER Theta
        A vector (of dimension (2x1)) of the slope and intercept
    parameters.

    .PARAMETER LearningRate
        A value of type [double] that signifies the magnitude of each
    step in the direction of the gradient.

    .PARAMETER Iterations
        A value of type [int] for the number of iterations to run
    gradient descent.
    
    .EXAMPLE
        $params = @{
            X = $X_matrix_augmented
            Y = $Y_vector
            Theta = $Theta
            LearningRate = 0.1
            Iterations = 100
        }
        $Theta, $_ = Run-GradientDescent @params
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]$X,
        [Parameter(Mandatory)]$Y,
        [Parameter(Mandatory)]$Theta,
        [Parameter(Mandatory)]$LearningRate,
        [Parameter(Mandatory)]$Iterations
    )

    $m = $Y.Count
    $CostHistory = New-Object double[] $Iterations

    for ($i=0; $i -lt $Iterations; $i++) {
        $predictions = $X * $Theta
        $gradient = (1/$m) * $X.Transpose() * ($predictions - $Y)
        $Theta = $Theta - $LearningRate * $gradient
        $CostHistory[$i] = Measure-CostFunction -X $X -Y $Y -Theta $Theta
    }
        
    return $Theta, $CostHistory
}
