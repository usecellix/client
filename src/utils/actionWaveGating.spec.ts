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
