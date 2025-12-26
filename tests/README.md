# Testing

## Requirements

- [Pytest-workflow](https://github.com/LUMC/pytest-workflow)
- [csvtk](https://github.com/shenwei356/csvtk) for processing of Achab Excel (path to exe hard-defined in `tests/wwwachab/run_and_specials.sh`)

## Directories organization

`tests` sub-dir contains 1 `test_my_workflow.yaml` file for each tested workflow or sub-workflow

Test datas for each workflow are expected to be under `tests/my_workflow/`

<br>

## Run tests

### General syntax

```bash
pytest \
	--verbose \
	--git-aware \
	tests/
```

Explanations (see [documentation](https://pytest-workflow.readthedocs.io/en/stable/) for details) :
* With `--basetemp=`, pytest-wf will run test inside `/scrath/<TEST_NAME>`
* Input JSON are configured so that everything lies inside `/scrath/<TEST_NAME>` (cromwell-exec + pipeline output)
* Due to `--keep-workflow-wd` option, all dirs created by pytest-wf must be be deleted *manually* afterwards
* Remove `--git-aware` if uncommited changes or outside git repo (otherwise error)


### Available tests

```bash
# Show Achab version:
pytest \
        --tag version \
        --verbose \
        --git-aware \
        tests/test_wwwachab.yaml

# Show Achab help:
pytest \
        --tag help \
        --verbose \
        --git-aware \
        tests/test_wwwachab.yaml

# Downsampled (only chr22) exome VCF with minimal Achab params:
pytest \
        --tag solo --tag exome --tag downsampled --tag minimal \
        --verbose \
        --git-aware \
        tests/test_wwwachab.yaml

# Downsampled (only chr22) exome VCF with Achab params used in production:
pytest \
        --tag solo --tag exome --tag downsampled --tag prod \
        --verbose \
        --git-aware \
        tests/test_wwwachab.yaml

# Full trio of exome (with params used in production):
pytest \
        --tag trio --tag exome --tag full --tag prod \
        --verbose \
        --git-aware \
        tests/test_wwwachab.yaml

```
