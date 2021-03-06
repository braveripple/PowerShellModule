Import-Module odbc
$connstr = "Driver={PostgreSQL Unicode(x64)};Server=localhost;Database=dvdrental;UID=postgres;PWD=postgres;Port=5432;"
try {
  $countSQL = "SELECT COUNT(*) FROM actor;"
  Get-OdbcData -C $connstr -Q $countSQL
  
  $conn = Get-OdbcConnection -C $connstr

  $tran = $conn.BeginTransaction()
  $sql = "INSERT INTO public.actor(actor_id, first_name, last_name, last_update) VALUES (2001, 'aaa', 'bbb', Now());"
  $rec = Invoke-OdbcQueryWithConnection -C $conn -Q $sql -T $tran
  $sql = "INSERT INTO public.actor(actor_id, first_name, last_name, last_update) VALUES (2002, 'ccc', 'ddd', Now());"
  $rec = Invoke-OdbcQueryWithConnection -C $conn -Q $sql -T $tran

  $tran.Commit();

} catch {
  Write-Error($_.Exception)
  try {
    Write-Host("Attempt to roll back the transaction.")
    $tran.Rollback();
  } catch {
    # Do nothing here; transaction is not active.
  }
} finally {
  try {
    Get-OdbcData -C $connstr -Q $countSQL
    $conn.Close();
  } catch {
    # Do nothing here; connection is not active.
  }
}