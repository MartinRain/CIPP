import { useCallback, useEffect, useState } from 'react'
import PropTypes from 'prop-types'
import {
  Alert,
  Button,
  CircularProgress,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  FormControlLabel,
  Radio,
  RadioGroup,
  Stack,
  TextField,
  Typography,
} from '@mui/material'
import { ApiGetCall, ApiPostCall } from '../../api/ApiCall'
import { usePermissions } from '../../hooks/use-permissions'

const parsePolicy = (data) => {
  const policy = data?.Results
  if (!policy || typeof policy !== 'object' || typeof policy.EnableAutoTrim !== 'boolean') {
    return null
  }
  return policy
}

export const CippEditGlobalVersionPolicyDialog = ({ open, onClose, tenantFilter }) => {
  const { checkPermissions } = usePermissions()
  const canRead = checkPermissions(['Sharepoint.Admin.Read', 'Sharepoint.Admin.ReadWrite'])
  const canWrite = checkPermissions(['Sharepoint.Admin.ReadWrite'])
  const validTenant = Boolean(tenantFilter && tenantFilter !== 'AllTenants')
  const [mode, setMode] = useState('automatic')
  const [major, setMajor] = useState('')
  const [days, setDays] = useState('')
  const [expectedPolicy, setExpectedPolicy] = useState(null)
  const [message, setMessage] = useState('')
  const [error, setError] = useState('')
  const policyQuery = ApiGetCall({
    url: '/api/ListSPOTenantVersionPolicy',
    data: { tenantFilter },
    queryKey: `SPOTenantVersionPolicy-${tenantFilter}`,
    waiting: open && canRead && validTenant,
    staleTime: 0,
    refetchOnMount: 'always',
    retry: 0,
  })
  // Refresh explicitly after a successful write; do not trigger a duplicate GET.
  const updatePolicy = ApiPostCall({})
  const policy = parsePolicy(policyQuery.data)

  const applyPolicy = useCallback((nextPolicy) => {
    setExpectedPolicy(nextPolicy)
    setMode(nextPolicy.EnableAutoTrim ? 'automatic' : 'manual')
    // Microsoft's automatic-policy sentinels are not valid manual limits.
    setMajor(nextPolicy.EnableAutoTrim ? '500' : String(nextPolicy.MajorVersionLimit ?? ''))
    setDays(nextPolicy.EnableAutoTrim ? '0' : String(nextPolicy.ExpireVersionsAfterDays ?? ''))
  }, [])

  useEffect(() => {
    if (!open) {
      setExpectedPolicy(null)
      setMessage('')
      setError('')
      return
    }
    if (!policy) return
    applyPolicy(policy)
  }, [open, policyQuery.data, tenantFilter, applyPolicy])

  const manual = mode === 'manual'
  const majorNumber = Number(major)
  const daysNumber = Number(days)
  const validManual =
    /^\d+$/.test(major) &&
    /^\d+$/.test(days) &&
    Number.isSafeInteger(majorNumber) &&
    majorNumber >= 100 &&
    majorNumber <= 50000 &&
    Number.isSafeInteger(daysNumber) &&
    (daysNumber === 0 || (daysNumber >= 30 && daysNumber <= 36500))
  const canSubmit = validTenant && canWrite && !!expectedPolicy && !policyQuery.isFetching &&
    !updatePolicy.isPending && (!manual || validManual)

  // React Query may retain the *same object reference* when a refresh returns
  // unchanged policy values. Restore the editor from the refetch result, rather
  // than depending on an effect that only observes data identity changes.
  const reloadPolicy = async () => {
    setExpectedPolicy(null)
    try {
      const response = await policyQuery.refetch()
      const refreshed = response?.isError ? null : parsePolicy(response?.data)
      if (!refreshed) {
        setError(
          response?.error?.response?.data?.Results ??
            response?.error?.message ??
            'Could not refresh the tenant version policy. Try Refresh again.'
        )
        return false
      }
      applyPolicy(refreshed)
      return true
    } catch (e) {
      setError(e?.response?.data?.Results ?? e?.message ?? 'Could not refresh the tenant version policy.')
      return false
    }
  }

  const save = async () => {
    if (!canSubmit) return
    setMessage('')
    setError('')
    const nextPolicy = {
      EnableAutoTrim: !manual,
      ...(manual
        ? { MajorVersionLimit: majorNumber, ExpireVersionsAfterDays: daysNumber }
        : {}),
    }
    const expected = {
      EnableAutoTrim: expectedPolicy.EnableAutoTrim,
      ...(expectedPolicy.EnableAutoTrim
        ? {}
        : {
            MajorVersionLimit: expectedPolicy.MajorVersionLimit,
            ExpireVersionsAfterDays: expectedPolicy.ExpireVersionsAfterDays,
          }),
    }
    try {
      const response = await updatePolicy.mutateAsync({
        url: '/api/ExecSPOTenantVersionPolicy',
        data: { tenantFilter, Policy: nextPolicy, ExpectedPolicy: expected },
      })
      const result = response?.data
      if (result?.Verified === false) {
        // State may still be propagating. Don't re-submit without reading actual state.
        setExpectedPolicy(null)
      } else {
        await reloadPolicy()
      }
      setMessage(result?.Results ?? 'SharePoint processed the request.')
    } catch (e) {
      setError(e?.response?.data?.Results ?? e?.message ?? 'Could not update the global version policy.')
      if (e?.response?.status === 409) setExpectedPolicy(null)
    }
  }

  const refresh = async () => {
    setMessage('')
    setError('')
    await reloadPolicy()
  }

  return (
    <Dialog open={open} onClose={updatePolicy.isPending ? undefined : onClose} fullWidth maxWidth="sm">
      <DialogTitle>Edit Global Version Policy</DialogTitle>
      <DialogContent dividers>
        <Stack spacing={2}>
          <Typography variant="body2" color="text.secondary">
            Tenant: {tenantFilter || 'None selected'}
          </Typography>
          <Alert severity="info">
            Changes to the tenant default affect new document libraries. They do not
            automatically update existing libraries or start a historical version cleanup job.
          </Alert>
          {!validTenant ? (
            <Alert severity="warning">Select a single tenant to edit its global policy.</Alert>
          ) : !canRead ? (
            <Alert severity="warning">SharePoint administrator read permission is required.</Alert>
          ) : policyQuery.isFetching && !expectedPolicy ? (
            <Stack direction="row" spacing={1} alignItems="center">
              <CircularProgress size={18} />
              <Typography variant="body2">Loading the current policy from SharePoint…</Typography>
            </Stack>
          ) : !policy ? (
            <Alert severity="error">
              {typeof policyQuery.data?.Results === 'string'
                ? policyQuery.data.Results
                : policyQuery.error?.response?.data?.Results ??
                  'Could not read the tenant policy. Check SharePoint app-only consent.'}
            </Alert>
          ) : (
            <>
              <RadioGroup
                value={mode}
                onChange={(e) => { setMode(e.target.value); setMessage(''); setError('') }}
                aria-label="Version policy mode"
              >
                <FormControlLabel
                  value="automatic"
                  control={<Radio />}
                  label="Automatic — Microsoft-managed intelligent version history"
                  disabled={!canWrite || !expectedPolicy}
                />
                <FormControlLabel
                  value="manual"
                  control={<Radio />}
                  label="Manual — fixed maximum and optional expiration"
                  disabled={!canWrite || !expectedPolicy}
                />
              </RadioGroup>
              {manual && (
                <Stack direction={{ xs: 'column', sm: 'row' }} spacing={2}>
                  <TextField
                    fullWidth
                    type="number"
                    label="Maximum major versions"
                    value={major}
                    onChange={(e) => setMajor(e.target.value)}
                    inputProps={{ min: 100, max: 50000, step: 1 }}
                    error={major !== '' && (!/^\d+$/.test(major) || majorNumber < 100 || majorNumber > 50000)}
                    helperText="100–50,000 major versions"
                    disabled={!canWrite || !expectedPolicy}
                  />
                  <TextField
                    fullWidth
                    type="number"
                    label="Expire after (days)"
                    value={days}
                    onChange={(e) => setDays(e.target.value)}
                    inputProps={{ min: 0, max: 36500, step: 1 }}
                    error={days !== '' && (!/^\d+$/.test(days) || !(daysNumber === 0 || (daysNumber >= 30 && daysNumber <= 36500)))}
                    helperText="0 = never; otherwise 30–36,500"
                    disabled={!canWrite || !expectedPolicy}
                  />
                </Stack>
              )}
              <Alert severity="warning">
                If a CIPP SPOVersionControl standard is configured with remediation enabled,
                its next run may override this change.
              </Alert>
            </>
          )}
          {error && <Alert severity="error">{error}</Alert>}
          {message && <Alert severity="success">{message}</Alert>}
        </Stack>
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose} color="inherit" disabled={updatePolicy.isPending}>
          Close
        </Button>
        <Button onClick={refresh} disabled={!validTenant || !canRead || policyQuery.isFetching || updatePolicy.isPending}>
          Refresh
        </Button>
        <Button variant="contained" onClick={save} disabled={!canSubmit}>
          {updatePolicy.isPending ? 'Saving…' : 'Save Policy'}
        </Button>
      </DialogActions>
    </Dialog>
  )
}

CippEditGlobalVersionPolicyDialog.propTypes = {
  open: PropTypes.bool.isRequired,
  onClose: PropTypes.func.isRequired,
  tenantFilter: PropTypes.string,
}

export default CippEditGlobalVersionPolicyDialog
