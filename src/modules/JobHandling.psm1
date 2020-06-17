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
            & $current.job
        }
    }
}