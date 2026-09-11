import { describe, expect, it } from 'vitest';
import {
  collectCascadeRejectIds,
  describeBlockedReason,
  isWaveDependencySatisfied,
} from './actionWaveGating';
import type { ActionBlock } from '@/types/conversationTurn';

function block(overrides: Partial<ActionBlock> & Pick<ActionBlock, 'id'>): ActionBlock {
  return {
    type: 'actions',
    actions: [],
    explanation: '',
    proposalStatus: 'pending',
    ...overrides,
  };
}

describe('isWaveDependencySatisfied', () => {
  it('is satisfied when the block has no dependency', () => {
    const b = block({ id: 'b1' });
    expect(isWaveDependencySatisfied(b, [])).toBe(true);
  });

  it('is not satisfied when the dependency is still pending', () => {
    const dep = block({ id: 'wave1', changeSetId: 'cs1', proposalStatus: 'pending' });
    const b = block({ id: 'wave2', dependsOnChangeSetId: 'cs1' });
    expect(isWaveDependencySatisfied(b, [dep, b])).toBe(false);
  });

  it('is not satisfied when the dependency was rejected', () => {
    const dep = block({ id: 'wave1', changeSetId: 'cs1', proposalStatus: 'rejected' });
    const b = block({ id: 'wave2', dependsOnChangeSetId: 'cs1' });
    expect(isWaveDependencySatisfied(b, [dep, b])).toBe(false);
  });

  it('is satisfied once the dependency is accepted', () => {
    const dep = block({ id: 'wave1', changeSetId: 'cs1', proposalStatus: 'accepted' });
    const b = block({ id: 'wave2', dependsOnChangeSetId: 'cs1' });
    expect(isWaveDependencySatisfied(b, [dep, b])).toBe(true);
  });

  it('is not satisfied when the dependency is missing entirely (e.g. history reload gap)', () => {
    const b = block({ id: 'wave2', dependsOnChangeSetId: 'cs-missing' });
    expect(isWaveDependencySatisfied(b, [b])).toBe(false);
  });

  // Live report (Sept 8, 2026): a 3-card progressive build where cards 1 and 2
  // both showed "Applied" but card 3 stayed stuck on "Accept the earlier step
  // first." A 3-link chain (card3 depends on card2, card2 depends on card1) was
  // untested — only 2-block chains existed above. This reproduces the exact
  // shape to check whether the gate itself is the bug, or whether the report is
  // explained by something else (e.g. stale UI, a race in when proposalStatus
  // actually flips to 'accepted').
  it('is satisfied for the third link of a chain once only its DIRECT dependency (not the root) is accepted', () => {
    const card1 = block({ id: 'card1', changeSetId: 'cs1', proposalStatus: 'accepted' });
    const card2 = block({
      id: 'card2',
      changeSetId: 'cs2',
      dependsOnChangeSetId: 'cs1',
      proposalStatus: 'accepted',
    });
    const card3 = block({ id: 'card3', changeSetId: 'cs3', dependsOnChangeSetId: 'cs2' });
    expect(isWaveDependencySatisfied(card3, [card1, card2, card3])).toBe(true);
  });

  it('is NOT satisfied for the third link when only the ROOT (card1) is accepted, not the direct parent (card2)', () => {
    const card1 = block({ id: 'card1', changeSetId: 'cs1', proposalStatus: 'accepted' });
    const card2 = block({
      id: 'card2',
      changeSetId: 'cs2',
      dependsOnChangeSetId: 'cs1',
      proposalStatus: 'pending', // NOT yet accepted
    });
    const card3 = block({ id: 'card3', changeSetId: 'cs3', dependsOnChangeSetId: 'cs2' });
    // This is the failure shape that WOULD explain the report: if card3's
    // dependsOnChangeSetId were ever wrongly seeded to point at card1 (the root)
    // instead of card2 (its direct predecessor), accepting card1 alone would
    // satisfy it even though card2 — the step actually between them — never
    // got accepted. Confirms the gate checks the DIRECT link only.
    expect(isWaveDependencySatisfied(card3, [card1, card2, card3])).toBe(false);
  });
});

describe('describeBlockedReason', () => {
  it('returns undefined for a block with no dependency', () => {
    expect(describeBlockedReason(block({ id: 'b1' }), [])).toBeUndefined();
  });

  it('explains a pending dependency', () => {
    const dep = block({ id: 'wave1', changeSetId: 'cs1', proposalStatus: 'pending' });
    const b = block({ id: 'wave2', dependsOnChangeSetId: 'cs1' });
    expect(describeBlockedReason(b, [dep, b])).toBe('Accept the earlier step first.');
  });

  it('explains a rejected dependency differently', () => {
    const dep = block({ id: 'wave1', changeSetId: 'cs1', proposalStatus: 'rejected' });
    const b = block({ id: 'wave2', dependsOnChangeSetId: 'cs1' });
    expect(describeBlockedReason(b, [dep, b])).toBe(
      'Skipped — an earlier required step was rejected.',
    );
  });

  it('returns undefined once the dependency is accepted', () => {
    const dep = block({ id: 'wave1', changeSetId: 'cs1', proposalStatus: 'accepted' });
    const b = block({ id: 'wave2', dependsOnChangeSetId: 'cs1' });
    expect(describeBlockedReason(b, [dep, b])).toBeUndefined();
  });
});

describe('selecting the block a summary-bar Accept should target', () => {
  // Regression for the silent no-op: App's findPendingActionBlock used to return
  // the first *pending* block regardless of its wave dependency. After wave 1 is
  // accepted, wave 2 is the first pending block — so the summary bar targeted it,
  // acceptActions refused it (dependency unmet is not the same as satisfied), and
  // the caller still cleared the preview. Net effect: nothing written, card gone,
  // no error shown. The selection must skip blocks that are not yet acceptable.
  const pick = (blocks: ActionBlock[]) =>
    blocks.find((b) => b.proposalStatus === 'pending' && isWaveDependencySatisfied(b, blocks));

  it('picks the first wave while it is still pending', () => {
    const wave1 = block({ id: 'wave1', changeSetId: 'cs1' });
    const wave2 = block({ id: 'wave2', changeSetId: 'cs2', dependsOnChangeSetId: 'cs1' });
    expect(pick([wave1, wave2])?.id).toBe('wave1');
  });

  it('picks the second wave only once the first is accepted', () => {
    const wave1 = block({ id: 'wave1', changeSetId: 'cs1', proposalStatus: 'accepted' });
    const wave2 = block({ id: 'wave2', changeSetId: 'cs2', dependsOnChangeSetId: 'cs1' });
    expect(pick([wave1, wave2])?.id).toBe('wave2');
  });

  it('picks nothing when the only pending block is still gated', () => {
    // The exact silent-no-op shape: wave1 pending (e.g. its own apply failed and
    // it was reset to pending), wave2 pending but gated behind it. Selecting
    // wave2 here is what produced an Accept that wrote nothing.
    const wave1 = block({ id: 'wave1', changeSetId: 'cs1', proposalStatus: 'rejected' });
    const wave2 = block({ id: 'wave2', changeSetId: 'cs2', dependsOnChangeSetId: 'cs1' });
    expect(pick([wave1, wave2])).toBeUndefined();
  });
});

describe('collectCascadeRejectIds', () => {
  it('includes just the rejected block when nothing depends on it', () => {
    const b1 = block({ id: 'wave1', changeSetId: 'cs1' });
    expect(collectCascadeRejectIds([b1], 'wave1')).toEqual(new Set(['wave1']));
  });

  it('cascades to a pending block that directly depends on the rejected one', () => {
    const wave1 = block({ id: 'wave1', changeSetId: 'cs1' });
    const wave2 = block({ id: 'wave2', dependsOnChangeSetId: 'cs1', changeSetId: 'cs2' });
    const ids = collectCascadeRejectIds([wave1, wave2], 'wave1');
    expect(ids).toEqual(new Set(['wave1', 'wave2']));
  });

  it('cascades transitively through a chain of dependencies', () => {
    const wave1 = block({ id: 'wave1', changeSetId: 'cs1' });
    const wave2 = block({ id: 'wave2', dependsOnChangeSetId: 'cs1', changeSetId: 'cs2' });
    const wave3 = block({ id: 'wave3', dependsOnChangeSetId: 'cs2', changeSetId: 'cs3' });
    const ids = collectCascadeRejectIds([wave1, wave2, wave3], 'wave1');
    expect(ids).toEqual(new Set(['wave1', 'wave2', 'wave3']));
  });

  it('does not cascade to a block that is already accepted', () => {
    const wave1 = block({ id: 'wave1', changeSetId: 'cs1' });
    const wave2 = block({
      id: 'wave2',
      dependsOnChangeSetId: 'cs1',
      changeSetId: 'cs2',
      proposalStatus: 'accepted',
    });
    // Rejecting wave1 after wave2 was already accepted should not un-accept wave2 —
    // it already happened. This should not occur in practice (wave2's Accept is
    // gated), but the traversal must not corrupt an already-applied wave.
    const ids = collectCascadeRejectIds([wave1, wave2], 'wave1');
    expect(ids).toEqual(new Set(['wave1']));
  });

  it('does not cascade to an unrelated block with a different dependency', () => {
    const wave1 = block({ id: 'wave1', changeSetId: 'cs1' });
    const other = block({ id: 'other', dependsOnChangeSetId: 'cs-unrelated', changeSetId: 'cs9' });
    const ids = collectCascadeRejectIds([wave1, other], 'wave1');
    expect(ids).toEqual(new Set(['wave1']));
  });
});
