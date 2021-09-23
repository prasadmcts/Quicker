protected async Task<ExpressionStarter<T>> GetAtSearchPredicate<T>(UserAuditTrialSearchDTO req, List<long> entityIds) where T : AuditTrailBase
        {
            await Task.Yield();
            var predicate = PredicateBuilder.New<T>(true);
            if (req.StartDate.HasValue && req.EndDate.HasValue)
            {
                predicate = predicate.And(p => p.CreatedOn >= req.StartDate && p.CreatedOn <= req.EndDate);
            }

            if (req.UserId.AnyOrDefault())
            {
                var userPredicate = PredicateBuilder.New<T>(true);
                foreach (var userId in req.UserId)
                {
                    userPredicate = userPredicate.Or(p => p.CreatedBy.Contains(userId));
                }
                predicate = predicate.And(userPredicate);
            }

            if (entityIds.AnyOrDefault())
            {
                var userPredicate = PredicateBuilder.New<T>(true);
                foreach (var entityId in entityIds.DistinctList())
                {
                    userPredicate = userPredicate.Or(p => p.EntityId == entityId);
                }
                predicate = predicate.And(userPredicate);
            }

            if (req.CreatedOn.HasValue)
            {
                var createdOnpredicate = PredicateBuilder.New<T>(true);
                createdOnpredicate = predicate.And(p => p.CreatedOn > req.CreatedOn);
                predicate = predicate.And(createdOnpredicate);
            }

            return predicate;
        }

===================

var predicate = await GetAtSearchPredicate<InspectionAuditTrail>(req, entityIds);
            return await _atCtx.InspectionsAuditTrail.AsNoTracking().AsExpandable().Where(predicate).ProjectTo<AuditTrailBaseReference>().ToListAsync();


using LinqKit;
