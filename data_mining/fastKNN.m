% The MIT License (MIT)
%
% Copyright (c) 2016 Markus Bergholz
% https://github.com/markuman/fastKNN

function [classified, dist] = KNNsearch(trained, data, k, distance)
    
    % loop the fastKNN for the whole dataset
    for i = 1:length(data(:,1))
        [_classified, _k, _dist] = fastKNN(trained, data(i,:), k, distance);

        classified(i) = _classified(1);
        dist(i)       = _dist(1);

    end

    classified=classified';
    dist=dist';

end % function KNNsearch

function [classified, k, dist, idx] = fastKNN(trained, data, k, distance)

    if (nargin <= 3)
        % Cosine Similarity
        % for p == 0, Cosine Similarity
        % Minkowski Distance
        % for p == 2, Minkowski becomes equal Euclidean
        % for p == 1, Minkowski becomes equal city block metric
        % for p ~= 1 && p ~= 2 -> Minkowski https://en.wikipedia.org/wiki/Minkowski_distance
        distance = 2;
    end

    % trained data has one more column as data, the class
    [dist, idx] = getDistance(trained, data, distance);

    if (nargin <= 2)
        % determine k value when no one is given
        % possible number of categories + 1
        k = numel( unique( trained(:,end) ) ) + 1;
    end
    
    classified  = mode(trained(idx(1:k), end));


end % function fastKNN

function [values, idx] = getDistance(x, y, p)

    if(p==0)    % Use Cosine Similarity
        for i = 1:length(x(:,1))
            _values(i) = 1-(x(i,1:end-1)*y')/(sqrt(sum(x(i,1:end-1).^2))*sqrt(sum(y.^2)'));
            _idx(i) = x(i,end);
        end
        
        % Transpose
        _values=_values';
        _idx=_idx';
        matrix = [_values, _idx];
        
        % Sort
        [values, idx] = sort(matrix,1);
        values=values(:,1);
        idx=idx(:,1);
        
    else        % Use Minkowski Distance
        [values, idx] = sort (sum ( abs( (x(:,1:end-1) - y(ones(size(x,1),1), :) ) .^p), 2) .^ (1 / p));
    end
end


