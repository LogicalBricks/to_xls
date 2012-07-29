module ToXls

  # Converts a hash with arrays keys to a hash with single (not arrays) keys, 
  # separating and merging the values correctly. If a key :all is specified, 
  # its value is merged into all other values when hashes.
  #
  # Example 1:
  #
  #     {
  #       [:a, :b] => {:c => :d}, 
  #       :e       => {:f => :g}, 
  #       :b       => {:h => :i},
  #       :all     => {:j => :k}
  #     }
  #
  # is converted to
  #
  #     {
  #       :a   => {:c => d, :j => :k},
  #       :b   => {:j => :k, :c => :d, :h => :i}, 
  #       :e   => {:h => :i, :f => :g},
  #       :all => {:j => :k}
  #     }
  #
  # Example 2:
  #
  #     {
  #       [:a, :b] => :c, 
  #       :d       => :e, 
  #       :b       => :f,
  #       :all     => :g
  #     }
  #
  # is converted to
  #
  #     {
  #       :a  => :c,
  #       :b  => :f, 
  #       :d  => :e,
  #       :all => :g
  #     }
  #
  class HashSimplifier
    def initialize hash
      @hash = hash
      @simple_hash = {}
    end

    def simple
      simplify
      apply_all
      @simple_hash
    end

    private

    def simplify
      @hash.each_pair do |k, v|
        [*k].each do |_k|
          if v.kind_of?(Hash)
            @simple_hash[_k] ||= {}
            @simple_hash[_k].merge!(v)
          else
            @simple_hash[_k] = v
          end
        end
      end
    end

    def apply_all
      all_value = @simple_hash[:all]
      if all_value
        @simple_hash.each {|k, v| v.merge!(all_value) if v.kind_of?(Hash) }
      end
    end

  end # class
end # module
