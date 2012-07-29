describe 'HashSimplifier' do
  context 'with hash as values' do
    describe 'simple keys' do
      it 'returns the same hash' do
        hash = {
          :a => {:some => :value}, 
          :b => {:other => :value}, 
          :c => {:multiple => :values, :in => :hash}
        }
        ToXls::HashSimplifier.new(hash).simple.should == {
          :a => {:some => :value}, 
          :b => {:other => :value}, 
          :c => {:multiple => :values, :in => :hash}
        }
      end
    end

    describe 'array keys' do
      it 'returns a hash with single (no array) keys' do
        hash = {
          [:a, :c] => {:some => :value}, 
          [:a, :b] => {:other => :value}, 
          :c => {:multiple => :values, :in => :hash}
        }
        ToXls::HashSimplifier.new(hash).simple.should == {
          :a => {:some  => :value, :other => :value}, 
          :b => {:other => :value}, 
          :c => {:some  => :value, :multiple => :values, :in => :hash}
        }
      end
    end

    describe 'hash with array keys and :all param' do
      it 'returns a hash with single (no array) keys' do
        hash = {
          [:a, :c] => {:some => :value}, 
          [:a, :b] => {:other => :value}, 
          :c => {:multiple => :values, :in => :hash}, 
          :all => {:apply => :to_all}
        }
        ToXls::HashSimplifier.new(hash).simple.should == {
          :a => {:some => :value, :other => :value, :apply => :to_all}, 
          :b => {:other => :value, :apply => :to_all}, 
          :c => {:some => :value, :multiple => :values, :in => :hash, :apply => :to_all},
          :all => {:apply => :to_all}
        }
      end
    end
  end

  context 'with simple values' do
    describe 'simple keys' do
      it 'returns the same hash' do
        hash = {
          :a => 'a',
          :b => 10,
          :c => :hash
        }
        ToXls::HashSimplifier.new(hash).simple.should == {
          :a => 'a',
          :b => 10,
          :c => :hash
        }
      end
    end

    describe 'array keys' do
      it 'returns an array with single (no array) key' do
        hash = {
          :a => 'a',
          [:a, :b] => 10,
          :c => :hash
        }
        ToXls::HashSimplifier.new(hash).simple.should == {
          :a => 10,
          :b => 10,
          :c => :hash
        }
      end
    end

    describe 'hash with array keys and :all param' do
      it 'returns a hash with single (no array) keys' do
        hash = {
          :a => 'a', 
          :b => 10, 
          :c => :hash,
          :all => 20
        }
        ToXls::HashSimplifier.new(hash).simple.should == {
          :a => 'a', 
          :b => 10, 
          :c => :hash,
          :all => 20
        }
      end
    end # describe
  end # context
end
