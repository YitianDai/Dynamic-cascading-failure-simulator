function [load_shedding,network_after] = Dynamic_Load_Shedding_Calculator(faulty_lines, faulty_nodes, target_load, amount_of_shedding)

    if ~exist('faulty_lines','var')
        faulty_lines = [];
    end
    
    if ~exist('faulty_nodes','var')
        faulty_nodes = [];
    end
    
    if ~exist('target_load','var')
        target_load = [3, 4, 7, 8, 12, 15, 16, 18, 20, 21, 23, 24, 25, 26, 27, 28, 29, 31, 39];
    end
    
    if ~exist('amount_of_shedding','var')
        amount_of_shedding = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    end    
    
    LS_balance = 0;
    LS_tripped = 0;
    LS_unsupplied = 0;
    
    network = case39_Digsilent;
    network.bus(:, 3:4) = network.bus(:, 3:4) * 1.0;
    load_initial = sum(network.bus(:, 3));
    network1 = network;
    network1.branch(faulty_lines, 11) = 0;
%     network2 = trip_nodes(network1, faulty_nodes);
    network1.gen(ismember(network1.gen(:, 1), faulty_nodes), :) = 0;
    network2 = network1;
    network3 = network2;
    [groups, isolated] = find_islands(network2);
    islands = [groups, num2cell(isolated)];
    fprintf('%d active island(s) found\n\n', size(islands, 2));
    
    for j = 1:size(islands, 2)

        % get the buses in this island
        island_nodes = network2.bus(islands{j}, 1);
        fprintf('Island %d: [', j);
        fprintf(repmat(' %d', 1, length(island_nodes)), island_nodes);
        fprintf(' ]\n');
        
        island = network2;
        island.bus(:, 2) = 4;
        island.bus(ismember(island.bus(:, 1), island_nodes), 2) = network2.bus(ismember(network2.bus(:, 1), island_nodes), 2);
        island.gen(~ismember(island.gen(:, 1), island_nodes), :) = 0; % set the parameters of generators that not in area to 0
        
        
        if island.gen(:, 9) == 0 % if there is no genrator in this island, trip all loads
            [network3, t] = trip_nodes(network3, island_nodes);
            fprintf('Island %d Tripped\n', j);
            LS_unsupplied = LS_unsupplied + sum(island.gen(:, 2))
        elseif size(island.bus(island.bus(:,2) ~= 4), 1) == 0 % if there is no active node in this area. trip all loads
            [network3, t] = trip_nodes(network3, island_nodes);
            fprintf('Island %d Tripped\n', j);
            LS_unsupplied = LS_unsupplied + sum(island.gen(:, 2))
        else
            fprintf('Island %d Survived\n', j);
        end
    end
    network_after = network3;
    load_after = sum(network_after.bus(:, 3));
    load_shedding = load_initial - load_after;
    for i = 1: length(target_load)
        if network_after.bus(target_load(i), 2)~= 4
            load_shedding = load_shedding + amount_of_shedding(i);
            
        end
    end  
end

% ,[3, 4, 7, 8, 12, 15, 16, 18, 20, 21, 23, 24, 25, 26, 27, 28, 29, 31, 39], [112.69999999999999, 137.5, 46.760000000000005, 104.4, 1.5, 112.0, 115.14999999999999, 55.3, 31.400000000000002, 95.89999999999999, 123.75, 108.01, 11.200000000000001, 13.9, 98.35, 20.6, 56.7, 1.8399999999999999, 220.8]

