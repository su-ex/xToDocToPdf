class ProgressHelper {
    [String] $activityName
    [String] $currentOperation
    [int] $totalOperations = 100
    [int] $i = 1

    ProgressHelper([String] $activityName) {
        $this.activityName = $activityName
        $this.currentOperation = "Start"
    }

    [void] setTotalOperations([int] $totalOperations) {
        $this.totalOperations = $totalOperations
    }

    [void] reset() {
        $this.i = 1
    }

    [void] update($operation) {
            $this.currentOperation = $operation
            Write-Progress -Activity $this.activityName -Status "Aufgabe $($this.i) von $($this.totalOperations)" -PercentComplete (($this.i/$this.totalOperations)*100) -CurrentOperation $this.currentOperation
            $this.i++
    }

    [void] error() {
        throw "Aufgabe `"$($this.currentOperation)`" konnte nicht ausgeführt werden."
    }

    [void] success() {
        Write-Progress -Activity $this.activityName -Status "Abgeschlossen" -PercentComplete 100
        Write-Host "Erfolgreich! :-)"
        Write-Progress -Activity $this.activityName -Completed
    }
}

Function ProgressHelper([String] $activityName) {
    return [ProgressHelper]::new($activityName)
}