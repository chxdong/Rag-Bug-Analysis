"""Read-only integrity checks and descriptive statistics for the released data."""

import csv
import hashlib
import json
import math
import re
from collections import Counter
from datetime import datetime
from pathlib import Path
from statistics import median

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
PROJECTS = {'infiniflow/ragflow': 'RAGFlow', 'Mintplex-Labs/anything-llm': 'AnythingLLM',
            'labring/FastGPT': 'FastGPT', 'Cinnamon/kotaemon': 'Kotaemon'}
EXPECTED_PROJECTS = {'RAGFlow': 160, 'AnythingLLM': 88, 'FastGPT': 60, 'Kotaemon': 25}
EXPECTED_SYMPTOMS = {'functional_failure': 201, 'crash': 111, 'build_failure': 11, 'performance_cost_issue': 10}
EXPECTED_ROOTS = {'data_processing_error': 81, 'logic_validation_error': 62, 'api_contract_error': 52,
                  'state_consistency_error': 46, 'dependency_environment_error': 39,
                  'resource_mismanagement': 26, 'security_vulnerability': 21, 'prompt_context_error': 6}
EXPECTED_DOMAIN = {'data_processing_error': 30, 'logic_validation_error': 11, 'api_contract_error': 11,
                   'state_consistency_error': 13, 'dependency_environment_error': 5,
                   'resource_mismanagement': 5, 'security_vulnerability': 0, 'prompt_context_error': 2}
EXPECTED_STAGES = {'knowledge_preparation': 110, 'ui_api_deployment': 87, 'workflow_agent_orchestration': 37,
                   'knowledge_base_management': 31, 'model_provider_integration': 26,
                   'retrieval_reranking': 24, 'response_generation': 18}
ALLOWED_FIXES = {'add_validation_or_guard', 'fix_data_processing_or_indexing', 'fix_retrieval_or_ranking',
                 'fix_prompt_or_generation', 'adapt_api_or_schema', 'fix_state_or_workflow_logic',
                 'fix_resource_handling', 'update_dependency_or_config', 'fix_ui_or_api_behavior'}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def read_csv(name):
    with (ROOT / 'data' / name).open(encoding='utf-8-sig', newline='') as stream:
        return list(csv.DictReader(stream))


def project(row):
    return PROJECTS['/'.join(row['issue_html_url'].split('/')[3:5])]


def unique_rows(rows, count):
    mapping = {r['issue_html_url']: r for r in rows}
    require(len(rows) == len(mapping) == count, f'Expected {count} unique issues')
    return mapping


def first_integer(value):
    match = re.match(r'\d+', str(value).replace(',', '').strip())
    require(match is not None, 'Invalid PR change count')
    return int(match.group())


def distributions(rows, field):
    by_project = {}
    for name, total in EXPECTED_PROJECTS.items():
        counts = Counter(r[field] for r in rows if project(r) == name)
        by_project[name] = {label: {'count': counts[label], 'percent': round(counts[label] / total * 100, 2)}
                            for label in sorted(set(r[field] for r in rows))}
    total = Counter(r[field] for r in rows)
    overall = {label: {'count': count, 'all_percent': round(count / len(rows) * 100, 2),
                        'equal_percent': round(sum(sum(r[field] == label for r in rows if project(r) == name) / n
                                                   for name, n in EXPECTED_PROJECTS.items()) / 4 * 100, 2)}
               for label, count in total.most_common()}
    return {'overall': overall, 'within_framework': by_project}


def verify_checksums():
    listed = set()
    for line in (ROOT / 'checksums.sha256').read_text(encoding='utf-8').splitlines():
        expected, name = line.split('  ', 1)
        path = (ROOT / name).resolve()
        require(path.is_relative_to(ROOT), 'Checksum path is outside the repository')
        require(path.is_file(), f'Missing file: {name}')
        require(hashlib.sha256(path.read_bytes()).hexdigest() == expected, f'Checksum mismatch: {name}')
        listed.add(name)
    actual = {p.relative_to(ROOT).as_posix() for p in ROOT.rglob('*') if p.is_file()
              and p.name != 'checksums.sha256' and '__pycache__' not in p.parts and '.git' not in p.parts}
    require(listed == actual, 'Files differ from the release manifest')


def main():
    verify_checksums()
    book = openpyxl.load_workbook(ROOT / 'BUG汇总.xlsx', read_only=True, data_only=True)
    values = book.active.iter_rows(values_only=True)
    headers = next(values)
    rows = [dict(zip(headers, row)) for row in values if any(v is not None for v in row)]
    book.close()
    base = unique_rows(rows, 333)
    require(Counter(project(r) for r in rows) == EXPECTED_PROJECTS, 'Framework counts differ')
    require(Counter(r['bug_run_symptom_type'] for r in rows) == EXPECTED_SYMPTOMS, 'Symptom counts differ')
    require(Counter(r['bug_root_cause_type'] for r in rows) == EXPECTED_ROOTS, 'Root counts differ')
    require(Counter(r['bug_stage'] for r in rows) == EXPECTED_STAGES, 'Stage counts differ')
    require(set(r['fix_strategy'] for r in rows) <= ALLOWED_FIXES, 'Unknown repair-operation label')
    require(all(r['issue_state'] == 'closed' for r in rows), 'Unexpected issue state')

    domain = read_csv('RAG根因二级标注.csv')
    require(set(unique_rows(domain, 333)) == set(base), 'Secondary labels do not cover the master table')
    require(Counter(r['rag_specific'] for r in domain) == {'1': 77, '0': 256}, 'Secondary label counts differ')
    for row in domain:
        source = base[row['issue_html_url']]
        require(row['pr_url'] == source['pr_url'] and row['bug_root_cause_type'] == source['bug_root_cause_type'], 'Secondary label key mismatch')
        require(row['framework'] == project(source), 'Secondary framework mismatch')
        require(row['rag_specific_label'] == ('RAG特有根因' if row['rag_specific'] == '1' else '通用根因'), 'Secondary label text mismatch')
    rag_counts = Counter(r['bug_root_cause_type'] for r in domain if r['rag_specific'] == '1')
    require(all(rag_counts[k] == v for k, v in EXPECTED_DOMAIN.items()), 'RAG-specific per-category counts differ')

    metrics = read_csv('RQ3_样本与指标.csv')
    require(set(unique_rows(metrics, 333)) == set(base), 'RQ3 keys differ')
    require(Counter(r['rq3_included'] for r in metrics) == {'1': 317, '0': 16}, 'RQ3 inclusion counts differ')
    for row in metrics:
        source = base[row['issue_html_url']]
        require(row['framework'] == project(source) and row['pr_url'] == source['pr_url'], 'RQ3 key mismatch')
        require(row['issue_created_at'] == source['issue_created_at'] and row['issue_closed_at'] == source['issue_closed_at'], 'RQ3 timestamp mismatch')
        duration = (datetime.fromisoformat(row['issue_closed_at']) - datetime.fromisoformat(row['issue_created_at'])).total_seconds() / 86400
        require(duration >= 0 and math.isclose(duration, float(row['open_duration_days']), abs_tol=1e-9), 'RQ3 duration mismatch')
        require(int(row['issue_comments']) == int(source['issue_comments']), 'RQ3 issue comments mismatch')
        require(int(row['pr_comments']) >= 0, 'Negative PR comment count')
        require(int(row['discussion_count']) == int(row['issue_comments']) + int(row['pr_comments']), 'RQ3 comment sum mismatch')
        require(int(row['changed_lines']) == first_integer(source['pr_changed_code_line']), 'RQ3 LOC mismatch')
        require(int(row['changed_files']) == first_integer(source['pr_changed_files']), 'RQ3 changed-files mismatch')
        require(bool(row['rq3_exclusion_reason']) == (row['rq3_included'] == '0'), 'RQ3 exclusion reason mismatch')
    included = [r for r in metrics if r['rq3_included'] == '1']
    medians = {k: median(float(r[k]) for r in included) for k in ['open_duration_days', 'issue_comments', 'pr_comments', 'discussion_count', 'changed_lines']}
    require(round(medians['open_duration_days'], 2) == 3.08 and medians['discussion_count'] == 3 and medians['changed_lines'] == 23, 'RQ3 summary differs')

    repairs = read_csv('RQ5_候选代码与语义审查.csv')
    require(set(unique_rows(repairs, 122)) <= set(base), 'RQ5 task outside the master dataset')
    for row in repairs:
        source = base[row['issue_html_url']]
        require(row['framework'] == project(source), 'RQ5 framework mismatch')
        for label in ['bug_stage', 'bug_run_symptom_type', 'bug_root_cause_type', 'fix_strategy']:
            require(row[label] == source[label], 'RQ5 primary label mismatch')
        for key in ['deepseek_semantic_accepted', 'gpt_semantic_accepted']:
            require(row[key] in {'0', '1'}, 'Invalid semantic-review value')
        require(bool(row['deepseek_candidate_code'].strip()) and bool(row['gpt_candidate_code'].strip()), 'Missing candidate output')
    accepted = {key: sum(int(r[key]) for r in repairs) for key in ['deepseek_semantic_accepted', 'gpt_semantic_accepted']}
    require(accepted == {'deepseek_semantic_accepted': 54, 'gpt_semantic_accepted': 51}, 'RQ5 acceptance counts differ')
    result = {'status': 'PASS', 'framework_counts': EXPECTED_PROJECTS,
              'symptoms': distributions(rows, 'bug_run_symptom_type'),
              'root_causes': distributions(rows, 'bug_root_cause_type'),
              'stage_counts': EXPECTED_STAGES,
              'rag_specific': {'count': 77, 'percent': round(77 / 333 * 100, 2), 'by_root': dict(rag_counts)},
              'general': {'count': 256, 'percent': round(256 / 333 * 100, 2)},
              'rq3': {'included': len(included), 'excluded': 16, 'medians': medians},
              'rq5': {'tasks': len(repairs), 'semantically_accepted': accepted,
                      'acceptance_percent': {k: round(v / 122 * 100, 2) for k, v in accepted.items()}}}
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
