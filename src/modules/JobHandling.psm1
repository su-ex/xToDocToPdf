class JobHandling {
    [String] $activityName
    [System.Collections.Queue] $jobs = @()

    JobHandling([String] $activityName) {
        $this.activityName = $activityName
    }

    [void] add([String]$operation, [ScriptBlock]$job) {
        $this.jobs.Enqueue([PSCustomObject]@{
            operation = $operation
            job = $job
        })
    }

    [void] run() {
        $total = $this.jobs.Count
        $i = 0
        while ($this.jobs.Count -gt 0) {
            $current = $this.jobs.Dequeue()
            Write-Progress -Activity $this.activityName -Status "Aufgabe $($i+1) von $total" -PercentComplete (($i++/$total)*100) -CurrentOperation $current.operation
            try {
                Write-Host $current.job
                $ret = (Invoke-Command -ScriptBlock $current.job -NoNewScope)
                Write-Host $ret
                if ($ret -ne $True) { throw }
            } catch {
                Write-Error "Aufgabe `"$($current.operation)`" konnte nicht ausgeführt werden."
                return
            }
        }
        Write-Progress -Activity $this.activityName -Status "Abgeschlossen" -PercentComplete 100
        Write-Host "Erfolgreich! :-)"
    }
}

Function JobHandling([String] $activityName) {
    return [JobHandling]::new($activityName)
}