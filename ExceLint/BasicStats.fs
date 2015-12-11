namespace ExceLint
    module BasicStats =
        
        /// <summary>Indicator function.</summary>
        /// <param name="b">index</param>
        /// <param name="t">threshold</param>
        /// <param name="xs">input vector</param>
        /// <returns>1 if true, 0 if false</returns>
        let indicator(b: int)(t: double)(xs: double[]) : double =
            if xs.[b] < t then 1.0 else 0.0

        /// <summary>Given a threshold value, computes the proportion of values < t.</summary>
        /// <param name="t">threshold</param>
        /// <param name="xs">input vector</param>
        /// <returns>a proportion between 0 and 1</returns>
        let cdf(t: double)(xs: double[]) : double =
            let inds = Array.map (fun i -> indicator i t xs) [| 0 .. (xs.Length - 1) |]
            (Array.sum inds) / System.Convert.ToDouble(xs.Length)

        /// <summary>For a given p value, find the value of x in X such that p <= Pr(x).</summary>
        /// <param name="p">p cumulative probability</param>
        /// <param name="xs">input vector</param>
        /// <returns>a value of x in X</returns>
        let cdfInverse(p: double)(xs: double[]) : double =
            assert (p >= 0.0)
            assert (p <= 1.0)

            // sort and zip with index
            let xs' = Array.sort xs

            // find the first index where cumPr > 
            let idx_opt = Array.tryFindIndex (fun x ->
                            (cdf x xs') > p
                          ) xs'

            match idx_opt with
            // if an index is found, take the one before
            | Some i -> if i = 0 then xs'.[0] else xs'.[i - 1]
            // if an index is not found, take the max element
            | None   -> xs'.[xs'.Length - 1]