using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;

namespace EasyDev.Expressions
{
        internal class ParameterReplacer : ExpressionVisitor
        {
                public ParameterExpression ParameterExpression { get; private set; }

                public ParameterReplacer(ParameterExpression paramExpr)
                {
                        this.ParameterExpression = paramExpr;
                }

                public Expression Replace(Expression expr)
                {
                        return this.Visit(expr);
                }

                protected override Expression VisitParameter(ParameterExpression p)
                {
                        return this.ParameterExpression;
                }
        }
}
