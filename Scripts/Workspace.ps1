
using namespace System.Management.Automation

class Workspace {

    [PathInfo]$location

    Workspace([PathInfo]$location = { Get-Location }) {
        $this.location = $location
    }
}
