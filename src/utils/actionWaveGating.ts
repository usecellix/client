import type { ActionBlock } from '@/types/conversationTurn';

/**
 * Staged accept waves (large multi-sheet builds split "create the sheets" from
 * "fill them in" — see conversation.service.ts's createActionWaveChangeSets):
 * a block with dependsOnChangeSetId set must not be accepted until the sibling
 * block carrying that changeSetId is itself 'accepted'.
 *
 * This logic previously lived inline, duplicated between TurnRenderer (for the
 * disabled-button explanation) and acceptActions (for the actual guard) — the
 * exact kind of two-implementations-of-one-rule drift risk this project keeps
 * finding bugs in. One source now, used by both.
 */

function findDependency(block: ActionBlock, siblings: ActionBlock[]): ActionBlock | undefined {
  if (!block.dependsOnChangeSetId) return undefined;
  return siblings.find((b) => b.changeSetId === block.dependsOnChangeSetId);
}

/** True when the block has no dependency, or its dependency has been accepted. */
export function isWaveDependencySatisfied(block: ActionBlock, siblings: ActionBlock[]): boolean {
  const dependency = findDependency(block, siblings);
  if (!block.dependsOnChangeSetId) return true;
  return dependency?.proposalStatus === 'accepted';
}

/** User-facing reason Accept is disabled, or undefined when it isn't blocked. */
export function describeBlockedReason(
  block: ActionBlock,
  siblings: ActionBlock[],
): string | undefined {
  if (!block.dependsOnChangeSetId) return undefined;
  const dependency = findDependency(block, siblings);
  if (dependency?.proposalStatus === 'accepted') return undefined;
  if (dependency?.proposalStatus === 'rejected') {
    return 'Skipped — an earlier required step was rejected.';
  }
  return 'Accept the earlier step first.';
}

/**
 * IDs of blocks that must be rejected alongside rejectedBlockId — every
 * pending block that depends on it, directly or transitively (its actions
 * targeted sheets/ranges the rejected wave would have created). Includes
 * rejectedBlockId itself.
 */
export function collectCascadeRejectIds(
  blocks: ActionBlock[],
  rejectedBlockId: string,
): Set<string> {
  const rejectedBlock = blocks.find((b) => b.id === rejectedBlockId);
  const cascadeIds = new Set<string>([rejectedBlockId]);
  let frontier = new Set<string>(rejectedBlock?.changeSetId ? [rejectedBlock.changeSetId] : []);

  while (frontier.size > 0) {
    const nextFrontier = new Set<string>();
    for (const block of blocks) {
      if (
        block.proposalStatus === 'pending' &&
        block.dependsOnChangeSetId &&
        frontier.has(block.dependsOnChangeSetId) &&
        !cascadeIds.has(block.id)
      ) {
        cascadeIds.add(block.id);
        if (block.changeSetId) nextFrontier.add(block.changeSetId);
      }
    }
    frontier = nextFrontier;
  }

  return cascadeIds;
}
