namespace DesktopAutomationAgent.ObjectRepository;

public sealed class ObjectReferenceResolver
{
    public ObjectResolutionResult Resolve(ObjectRepositorySnapshot snapshot, string reference)
    {
        ArgumentNullException.ThrowIfNull(snapshot);

        if (string.IsNullOrWhiteSpace(reference))
        {
            return new ObjectResolutionResult
            {
                Reference = reference ?? string.Empty,
                Errors = ["Object reference is required (format: pageId.elementId)."]
            };
        }

        var trimmed = reference.Trim();
        var dotIndex = trimmed.IndexOf('.');
        if (dotIndex <= 0 || dotIndex == trimmed.Length - 1)
        {
            return new ObjectResolutionResult
            {
                Reference = trimmed,
                Errors = [$"Object reference '{trimmed}' must use the form pageId.elementId."]
            };
        }

        var pageId = trimmed[..dotIndex];
        var elementId = trimmed[(dotIndex + 1)..];

        if (!snapshot.Pages.TryGetValue(pageId, out var page))
        {
            return new ObjectResolutionResult
            {
                Reference = trimmed,
                PageId = pageId,
                ElementId = elementId,
                Errors = [$"Page '{pageId}' was not found in the object repository."]
            };
        }

        if (!string.Equals(page.State, "active", StringComparison.Ordinal))
        {
            return new ObjectResolutionResult
            {
                Reference = trimmed,
                PageId = pageId,
                ElementId = elementId,
                PageName = page.Name,
                Errors = [$"Page '{pageId}' is not active (state='{page.State}')."]
            };
        }

        if (page.Elements is null || !page.Elements.TryGetValue(elementId, out var element))
        {
            return new ObjectResolutionResult
            {
                Reference = trimmed,
                PageId = pageId,
                ElementId = elementId,
                PageName = page.Name,
                Errors = [$"Element '{elementId}' was not found on page '{pageId}'."]
            };
        }

        if (element.Locator is null)
        {
            return new ObjectResolutionResult
            {
                Reference = trimmed,
                PageId = pageId,
                ElementId = elementId,
                PageName = page.Name,
                ElementDescription = element.Description,
                Errors = [$"Element '{pageId}.{elementId}' does not define a locator."]
            };
        }

        var locatorValidation = ObjectLocatorValidator.Validate(element.Locator, $"{pageId}.{elementId}.locator");
        if (!locatorValidation.IsValid)
        {
            return new ObjectResolutionResult
            {
                Reference = trimmed,
                PageId = pageId,
                ElementId = elementId,
                PageName = page.Name,
                ElementDescription = element.Description,
                Locator = element.Locator,
                Errors = locatorValidation.Errors,
                Warnings = locatorValidation.Warnings
            };
        }

        return new ObjectResolutionResult
        {
            Reference = trimmed,
            PageId = pageId,
            ElementId = elementId,
            PageName = page.Name,
            ElementDescription = element.Description,
            Locator = element.Locator,
            Warnings = locatorValidation.Warnings
        };
    }
}
