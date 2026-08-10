using DesktopAutomationAgent.Plans;

namespace DesktopAutomationAgent.ObjectRepository;

public sealed class PlanObjectRepositoryIntegrator
{
    private readonly ObjectRepositoryReader _repositoryReader;
    private readonly PlanObjectReferenceExpander _expander;

    public PlanObjectRepositoryIntegrator(
        ObjectRepositoryReader repositoryReader,
        PlanObjectReferenceExpander expander)
    {
        _repositoryReader = repositoryReader;
        _expander = expander;
    }

    public PlanValidationResult Integrate(PlanValidationResult validation)
    {
        ArgumentNullException.ThrowIfNull(validation);

        if (!validation.IsValid || validation.Plan is null)
            return validation;

        var plan = validation.Plan;
        if (!RequiresIntegration(plan))
            return validation;

        var errors = validation.Errors.ToList();
        var warnings = validation.Warnings?.ToList() ?? [];

        ObjectRepositorySnapshot? snapshot = null;
        if (!string.IsNullOrWhiteSpace(plan.ObjectRepository))
        {
            var repositoryResult = _repositoryReader.Read(plan.ObjectRepository);
            warnings.AddRange(repositoryResult.Warnings);
            if (!repositoryResult.IsValid || repositoryResult.Snapshot is null)
            {
                errors.AddRange(repositoryResult.Errors);
                return Rebuild(validation, errors, warnings, null, null, attachRepositoryAudit: false);
            }

            snapshot = repositoryResult.Snapshot;
        }

        if (snapshot is null)
        {
            errors.Add($"{validation.PlanPath}: object repository is required when the plan contains $objectRef markers.");
            return Rebuild(validation, errors, warnings, null, null, attachRepositoryAudit: false);
        }

        var expansion = _expander.Expand(plan, snapshot, validation.PlanPath);
        warnings.AddRange(expansion.Warnings);
        if (!expansion.Success)
        {
            errors.AddRange(expansion.Errors);
            return Rebuild(validation, errors, warnings, expansion, null, attachRepositoryAudit: false);
        }

        return Rebuild(
            validation,
            errors,
            warnings,
            expansion,
            plan,
            attachRepositoryAudit: expansion.ResolvedObjectReferences.Count > 0);
    }

    private static bool RequiresIntegration(PlanManifest plan)
    {
        if (!string.IsNullOrWhiteSpace(plan.ObjectRepository))
            return true;

        return ContainsObjectRef(plan.Steps) || ContainsObjectRef(plan.OnFailureSteps);
    }

    private static bool ContainsObjectRef(List<PlanStep>? steps)
    {
        if (steps is null)
            return false;

        foreach (var step in steps)
        {
            if (step.Arguments is null)
                continue;

            foreach (var value in step.Arguments.Values)
            {
                if (ContainsObjectRef(value))
                    return true;
            }
        }

        return false;
    }

    private static bool ContainsObjectRef(System.Text.Json.JsonElement element)
    {
        switch (element.ValueKind)
        {
            case System.Text.Json.JsonValueKind.Object:
                if (element.TryGetProperty("$objectRef", out _))
                    return true;

                foreach (var property in element.EnumerateObject())
                {
                    if (ContainsObjectRef(property.Value))
                        return true;
                }

                break;

            case System.Text.Json.JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                {
                    if (ContainsObjectRef(item))
                        return true;
                }

                break;
        }

        return false;
    }

    private static PlanValidationResult Rebuild(
        PlanValidationResult validation,
        List<string> errors,
        List<string> warnings,
        PlanObjectReferenceExpansionResult? expansion,
        PlanManifest? plan,
        bool attachRepositoryAudit) =>
        new()
        {
            PlanPath = validation.PlanPath,
            PlanId = validation.PlanId,
            Name = validation.Name,
            StepCount = validation.StepCount,
            OnFailureStepCount = validation.OnFailureStepCount,
            CleanupStepCount = validation.CleanupStepCount,
            TotalStepCount = validation.TotalStepCount,
            Sha256 = validation.Sha256,
            Errors = errors,
            Warnings = warnings,
            Plan = errors.Count == 0 ? (plan ?? validation.Plan) : null,
            ObjectRepositoryPath = attachRepositoryAudit ? expansion?.RepositoryPath : null,
            ObjectRepositoryId = attachRepositoryAudit ? expansion?.RepositoryId : null,
            ObjectRepositorySha256 = attachRepositoryAudit ? expansion?.RepositorySha256 : null,
            ResolvedObjectReferences = attachRepositoryAudit ? expansion?.ResolvedObjectReferences : null
        };
}
