"use strict";

import powerbi from "powerbi-visuals-api";
import {
    BasicFilter,
    FilterType,
    IBasicFilter,
    IFilterTarget,
    ITupleFilter,
    TupleFilter
} from "powerbi-models";
import { FieldModel, FieldState, ValueItem, emptyState, isChecked, sameTarget, toFilterValue, valueKey } from "./filters";

import PrimitiveValue = powerbi.PrimitiveValue;

const SEP = "\u001f";

export interface TreeNode {
    /** Caminho completo (valores de todos os níveis até aqui): Ano 2024 › Mês 3 ≠ Ano 2025 › Mês 3. */
    id: string;
    level: number;
    item: ValueItem;
    children: TreeNode[];
    /** Ids das folhas (último nível) abaixo deste nó. */
    leafKeys: string[];
}

/** Colunas de um mesmo espaço "Filtro aninhado N", mostradas como uma árvore. */
export interface NestedGroup {
    levels: FieldModel[];
    roots: TreeNode[];
    nodes: Map<string, TreeNode>;
    /** Valores de cada nível no caminho de cada folha. */
    leafPath: Map<string, PrimitiveValue[]>;
}

function childId(parent: TreeNode | undefined, key: string): string {
    return parent ? parent.id + SEP + key : key;
}

export function buildTree(levels: FieldModel[], columns: PrimitiveValue[][]): NestedGroup {
    const nodes = new Map<string, TreeNode>();
    const roots: TreeNode[] = [];
    const leafPath = new Map<string, PrimitiveValue[]>();
    const lookups = levels.map(f => new Map(f.values.map(v => [v.key, v])));
    const depth = levels.length;
    const rows = columns[0]?.length || 0;

    for (let r = 0; r < rows; r++) {
        let parent: TreeNode | undefined;
        for (let l = 0; l < depth; l++) {
            const raw = columns[l][r];
            const key = valueKey(raw);
            const id = childId(parent, key);
            let node = nodes.get(id);
            if (!node) {
                node = { id, level: l, item: lookups[l].get(key) || { key, raw, label: String(raw) }, children: [], leafKeys: [] };
                nodes.set(id, node);
                (parent ? parent.children : roots).push(node);
                if (l === depth - 1) {
                    leafPath.set(id, columns.map(c => c[r]));
                }
            }
            parent = node;
        }
    }

    const finish = (list: TreeNode[]): void => {
        for (const node of list) {
            if (node.children.length) {
                finish(node.children);
                node.leafKeys = node.children.flatMap(c => c.leafKeys);
            } else {
                node.leafKeys = [node.id];
            }
        }
    };
    finish(roots);

    return { levels, roots, nodes, leafPath };
}

function fromFilterValue(value: unknown, level: FieldModel): PrimitiveValue {
    if (level.kind === "date" && typeof value === "string") {
        return new Date(value);
    }
    return value as PrimitiveValue;
}

/** Nós mais altos que estão totalmente marcados (resumo da seleção). */
export function coverNodes(nodes: TreeNode[], selected: { has(key: string): boolean }): TreeNode[] {
    const result: TreeNode[] = [];
    for (const node of nodes) {
        const count = node.leafKeys.filter(k => selected.has(k)).length;
        if (count === 0) {
            continue;
        }
        if (count === node.leafKeys.length) {
            result.push(node);
        } else {
            result.push(...coverNodes(node.children, selected));
        }
    }
    return result;
}

/** Nós de um nível com o mesmo valor (ex.: o mês 3 aparece em 2024 e em 2025). */
function nodesWithValue(group: NestedGroup, level: number, key: string): TreeNode[] {
    const result: TreeNode[] = [];
    group.nodes.forEach(node => {
        if (node.level === level && node.item.key === key) {
            result.push(node);
        }
    });
    return result;
}

/** Filtro de várias colunas (tupla) com o caminho completo das folhas indicadas. */
function tupleFilter(group: NestedGroup, ids: string[]): ITupleFilter {
    const values = ids
        .map(id => group.leafPath.get(id))
        .filter(path => !!path)
        .map(path => path.map(raw => ({ value: toFilterValue(raw) })));
    return new TupleFilter(group.levels.map(l => l.target), "In", values).toJSON();
}

/**
 * Filtros da árvore. Se só há itens inteiros do 1º nível marcados, usa um filtro
 * simples nessa coluna; caso contrário, um filtro de várias colunas (tupla) com
 * o caminho completo de cada folha marcada, que é exato mesmo quando o mesmo
 * valor aparece em vários pais (ex.: o mês 3 de 2024 e o mês 3 de 2025).
 *
 * No modo "tudo exceto" (depois de "Selecionar tudo"), cada item desmarcado vira um
 * "NotIn" na coluna do seu nível (um filtro por nível; o Power BI combina com E).
 * Assim valores que não aparecem na lista agora (ex.: só existem no ano anterior)
 * continuam valendo. Se algum item desmarcado tem um valor que se repete em outro pai
 * (Mar de 2025, mas não Mar de 2024), "NotIn" na coluna não é exato: aí volta para a tupla.
 */
export function buildTreeFilters(group: NestedGroup, state: FieldState, singleSelect: boolean): Array<IBasicFilter | ITupleFilter> {
    if (state.invert) {
        if (state.excluded.size === 0) {
            return [];
        }
        const cover = coverNodes(group.roots, state.excluded);
        if (cover.every(n => n.level === 0 || nodesWithValue(group, n.level, n.item.key).length === 1)) {
            const filters: IBasicFilter[] = [];
            group.levels.forEach((level, l) => {
                const values = cover.filter(n => n.level === l).map(n => toFilterValue(n.item.raw));
                if (values.length) {
                    filters.push(new BasicFilter(level.target, "NotIn", values).toJSON());
                }
            });
            return filters;
        }
        const checked = Array.from(group.leafPath.keys()).filter(id => isChecked(state, id));
        return checked.length ? [tupleFilter(group, checked)] : [];
    }

    if (state.selected.size === 0) {
        return [];
    }
    const cover = coverNodes(group.roots, state.selected);
    if (cover.every(n => n.level === 0)) {
        const filter = new BasicFilter(group.levels[0].target, "In", cover.map(n => toFilterValue(n.item.raw)));
        filter.requireSingleSelection = singleSelect && cover.length === 1;
        return [filter.toJSON()];
    }
    return [tupleFilter(group, Array.from(state.selected.keys()))];
}

function sameTargets(a: IFilterTarget[], b: IFilterTarget[]): boolean {
    return a.length === b.length && a.every((t, i) => sameTarget(t, b[i]));
}

/** Reconstrói a seleção da árvore a partir dos filtros salvos no relatório. */
export function treeStateFromFilters(group: NestedGroup, filters: Array<{ filterType?: number; target?: unknown }>): FieldState {
    const state = emptyState();
    const targets = group.levels.map(l => l.target);

    for (const filter of filters) {
        if (!filter || !filter.target) {
            continue;
        }
        if (filter.filterType === FilterType.Tuple && Array.isArray(filter.target)
            && sameTargets(filter.target as IFilterTarget[], targets)) {
            for (const tuple of (filter as ITupleFilter).values) {
                let id = "";
                tuple.forEach((element, l) => {
                    const key = valueKey(fromFilterValue(element.value, group.levels[l]));
                    id = l === 0 ? key : id + SEP + key;
                });
                if (group.leafPath.has(id)) {
                    state.selected.set(id, null);
                }
            }
        } else if (filter.filterType === FilterType.Basic && !Array.isArray(filter.target)) {
            const basic = filter as IBasicFilter;
            const level = targets.findIndex(t => sameTarget(filter.target as IFilterTarget, t));
            if (basic.operator === "NotIn" && level >= 0) {
                // "Tudo exceto": um filtro por nível com os itens desmarcados.
                state.invert = true;
                for (const value of basic.values) {
                    const key = valueKey(fromFilterValue(value, group.levels[level]));
                    nodesWithValue(group, level, key).forEach(node => node.leafKeys.forEach(k => state.excluded.set(k, null)));
                }
            } else if (level === 0) {
                for (const value of basic.values) {
                    const root = group.nodes.get(valueKey(fromFilterValue(value, group.levels[0])));
                    root?.leafKeys.forEach(k => state.selected.set(k, null));
                }
            }
        }
    }
    if (state.invert) {
        // Os dois modos não se misturam: o "NotIn" já diz tudo o que está marcado.
        state.selected.clear();
    }
    return state;
}
