import React from 'react'
import { fireEvent, render, screen, waitFor } from '@testing-library/react'
import { CippEditGlobalVersionPolicyDialog } from '../../../src/components/CippComponents/CippEditGlobalVersionPolicyDialog'
import { ApiGetCall, ApiPostCall } from '../../../src/api/ApiCall'

vi.mock('../../../src/api/ApiCall', () => ({
  ApiGetCall: vi.fn(),
  ApiPostCall: vi.fn(),
}))
vi.mock('../../../src/hooks/use-permissions', () => ({
  usePermissions: () => ({
    checkPermissions: (roles) =>
      roles.includes('Sharepoint.Admin.Read') || roles.includes('Sharepoint.Admin.ReadWrite'),
  }),
}))

describe('CippEditGlobalVersionPolicyDialog', () => {
  let mutateAsync
  let refetch

  beforeEach(() => {
    vi.clearAllMocks()
    mutateAsync = vi.fn().mockResolvedValue({
      data: { Results: 'Global policy updated.', Verified: true, Changed: true },
    })
    refetch = vi.fn().mockResolvedValue({})
    ApiGetCall.mockReturnValue({
      data: { Results: { EnableAutoTrim: false, MajorVersionLimit: 500, ExpireVersionsAfterDays: 0 } },
      isFetching: false,
      refetch,
    })
    ApiPostCall.mockReturnValue({ mutateAsync, isPending: false })
  })

  it('loads the tenant policy only while the editor is open', () => {
    const { rerender } = render(
      <CippEditGlobalVersionPolicyDialog open={false} onClose={() => {}} tenantFilter="contoso.onmicrosoft.com" />
    )
    expect(ApiGetCall).toHaveBeenCalledWith(
      expect.objectContaining({ waiting: false, data: { tenantFilter: 'contoso.onmicrosoft.com' } })
    )
    rerender(
      <CippEditGlobalVersionPolicyDialog open onClose={() => {}} tenantFilter="contoso.onmicrosoft.com" />
    )
    expect(ApiGetCall).toHaveBeenCalledWith(
      expect.objectContaining({ waiting: true })
    )
    expect(screen.getByDisplayValue('500')).toBeInTheDocument()
    expect(screen.getByDisplayValue('0')).toBeInTheDocument()
  })

  it('sends the original policy as ExpectedPolicy to prevent stale updates', async () => {
    render(
      <CippEditGlobalVersionPolicyDialog open onClose={() => {}} tenantFilter="contoso.onmicrosoft.com" />
    )
    fireEvent.click(screen.getByLabelText(/Automatic — Microsoft-managed/i))
    fireEvent.click(screen.getByRole('button', { name: 'Save Policy' }))
    await waitFor(() => expect(mutateAsync).toHaveBeenCalledTimes(1))
    expect(mutateAsync.mock.calls[0][0]).toEqual({
      url: '/api/ExecSPOTenantVersionPolicy',
      data: {
        tenantFilter: 'contoso.onmicrosoft.com',
        Policy: { EnableAutoTrim: true },
        ExpectedPolicy: { EnableAutoTrim: false, MajorVersionLimit: 500, ExpireVersionsAfterDays: 0 },
      },
    })
  })

  it('blocks invalid manual limits before posting', () => {
    render(
      <CippEditGlobalVersionPolicyDialog open onClose={() => {}} tenantFilter="contoso.onmicrosoft.com" />
    )
    fireEvent.change(screen.getByLabelText('Maximum major versions'), { target: { value: '50' } })
    expect(screen.getByRole('button', { name: 'Save Policy' })).toBeDisabled()
    expect(mutateAsync).not.toHaveBeenCalled()
  })

  it('keeps the controls enabled after Refresh returns the same cached policy object', async () => {
    const unchanged = {
      Results: { EnableAutoTrim: false, MajorVersionLimit: 500, ExpireVersionsAfterDays: 0 },
    }
    ApiGetCall.mockReturnValue({
      data: unchanged,
      isFetching: false,
      refetch: vi.fn().mockResolvedValue({ data: unchanged, isError: false }),
    })
    render(
      <CippEditGlobalVersionPolicyDialog open onClose={() => {}} tenantFilter="contoso.onmicrosoft.com" />
    )
    const save = screen.getByRole('button', { name: 'Save Policy' })
    expect(save).toBeEnabled()
    fireEvent.click(screen.getByRole('button', { name: 'Refresh' }))
    await waitFor(() => expect(save).toBeEnabled())
    expect(screen.getByDisplayValue('500')).toBeEnabled()
    expect(screen.getByDisplayValue('0')).toBeEnabled()
    expect(screen.getByLabelText(/Automatic — Microsoft-managed/i)).toBeEnabled()
  })

  it('restores changed values returned by Refresh and uses them as the expected policy', async () => {
    const original = {
      Results: { EnableAutoTrim: false, MajorVersionLimit: 500, ExpireVersionsAfterDays: 0 },
    }
    const refreshed = {
      Results: { EnableAutoTrim: true, MajorVersionLimit: -1, ExpireVersionsAfterDays: -1 },
    }
    const refreshMock = vi.fn().mockResolvedValue({ data: refreshed, isError: false })
    ApiGetCall.mockReturnValue({ data: original, isFetching: false, refetch: refreshMock })
    render(
      <CippEditGlobalVersionPolicyDialog open onClose={() => {}} tenantFilter="contoso.onmicrosoft.com" />
    )
    fireEvent.click(screen.getByRole('button', { name: 'Refresh' }))
    await waitFor(() => expect(screen.getByLabelText(/Automatic — Microsoft-managed/i)).toBeChecked())
    expect(screen.getByRole('button', { name: 'Save Policy' })).toBeEnabled()
    fireEvent.click(screen.getByLabelText(/Manual — fixed maximum/i))
    fireEvent.click(screen.getByRole('button', { name: 'Save Policy' }))
    await waitFor(() => expect(mutateAsync).toHaveBeenCalledTimes(1))
    expect(mutateAsync.mock.calls[0][0].data.ExpectedPolicy).toEqual({ EnableAutoTrim: true })
  })

  it('shows a refresh error and keeps saving disabled when the read fails', async () => {
    ApiGetCall.mockReturnValue({
      data: { Results: { EnableAutoTrim: false, MajorVersionLimit: 500, ExpireVersionsAfterDays: 0 } },
      isFetching: false,
      refetch: vi.fn().mockResolvedValue({
        isError: true,
        error: { message: 'SharePoint unavailable' },
      }),
    })
    render(
      <CippEditGlobalVersionPolicyDialog open onClose={() => {}} tenantFilter="contoso.onmicrosoft.com" />
    )
    fireEvent.click(screen.getByRole('button', { name: 'Refresh' }))
    expect(await screen.findByText('SharePoint unavailable')).toBeInTheDocument()
    expect(screen.getByRole('button', { name: 'Save Policy' })).toBeDisabled()
    expect(mutateAsync).not.toHaveBeenCalled()
  })

  it('refuses the all-tenant context', () => {
    render(
      <CippEditGlobalVersionPolicyDialog open onClose={() => {}} tenantFilter="AllTenants" />
    )
    const save = screen.getByRole('button', { name: 'Save Policy' })
    expect(save).toBeDisabled()
    fireEvent.click(save)
    expect(mutateAsync).not.toHaveBeenCalled()
    expect(ApiGetCall).toHaveBeenCalledWith(expect.objectContaining({ waiting: false }))
  })
})
