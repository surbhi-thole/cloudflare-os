import { useEffect, useState } from 'react'
import { Badge, Tooltip, useKumoToastManager } from '@cloudflare/kumo'
import { ArrowClockwise, Camera, Plug } from '@phosphor-icons/react'
import { WorkshopIconButton } from './components/WorkshopControls'
import { formatRelativeTime } from './Activity'
import { formatFullTimestamp } from './utils/formatTimestamp'

type Props = {
  refreshedAt: Date
  onRefresh: () => Promise<void>
  // Asks to verify the collaborator's accounts and open the live gadget in place of the snapshot.
  onVerify: () => void
  // Whether that verification is under way.
  verifying: boolean
}

/**
 * Marks the gadget on screen as a static snapshot (see Overseer.getGadgetSnapshot), saying when it
 * was taken, with controls to verify for the live gadget and to take a new snapshot.
 */
export const GadgetSnapshotStatus = ({ refreshedAt, onRefresh, onVerify, verifying }: Props) => {
  const [refreshing, setRefreshing] = useState(false)
  const toasts = useKumoToastManager()
  // Re-render each minute so the relative age stays current.
  const [, setTick] = useState(0)
  useEffect(() => {
    const timer = setInterval(() => setTick(tick => tick + 1), 60_000)
    return () => clearInterval(timer)
  }, [])

  const refresh = async () => {
    setRefreshing(true)
    try {
      await onRefresh()
    } catch (error) {
      console.error('Failed to refresh gadget snapshot:', error)
      toasts.add({ title: 'Failed to refresh the snapshot', variant: 'error' })
    } finally {
      setRefreshing(false)
    }
  }

  return (
    <div className="flex min-w-0 items-center gap-1.5">
      <Badge variant="secondary" icon={<Camera size={12} />}>Snapshot</Badge>
      <span title={formatFullTimestamp(refreshedAt)} className="hidden whitespace-nowrap text-xs text-kumo-subtle lg:inline">
        updated {formatRelativeTime(refreshedAt)}
      </span>
      <Tooltip content="Verify for full access" asChild>
        <WorkshopIconButton aria-label="Verify for full access" disabled={verifying} onClick={onVerify}>
          <Plug size={17} />
        </WorkshopIconButton>
      </Tooltip>
      <Tooltip content="Refresh snapshot" asChild>
        <WorkshopIconButton aria-label="Refresh snapshot" disabled={refreshing} onClick={refresh}>
          <ArrowClockwise size={17} className={refreshing ? 'animate-spin' : undefined} />
        </WorkshopIconButton>
      </Tooltip>
    </div>
  )
}
