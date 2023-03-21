#!/usr/bin/env ruby
require 'rubyXL'
require 'pg'

class Item
    def initialize(item_id) 
        @item_id = item_id
        @labels = {}
    end

    def add_label(label, value)
        @labels[label] = value
    end

    def get_label(label)
        tmp = @labels[label]
        if tmp == nil
            return "NULL"
        end
        if tmp.instance_of? String
            if tmp == "YES"
                return true
            end
            if tmp == "NO"
                return false
            end
            return  "'" + tmp + "'";
        else     
            return tmp
        end
    end

    def get_id()
        @item_id
    end

    def print()
        printf("\tItem %i:\n", @item_id)
        @labels.each { |label, value|
            printf("\t\t%s: %s\n", label, value)
        }
    end
end

class Package
    def initialize(package_id) 
        @package_id = package_id
        @items = []
    end
    
    def add_or_update_item(item_id, label, value)
        item = @items.find {|itm| itm.get_id == item_id}
        if (item == nil)
            item = Item.new(item_id)
            @items << item
        end
        item.add_label(label, value)
    end

    def get_id()
        @package_id
    end

    def get_items()
        @items
    end

    def print()
        printf("Package %i:\n", @package_id)
        @items.each { |item|
            item.print()
        }
    end
end

class Orders
    def initialize()
        @packages = []
    end

    def init_from_xlsx(filePath)
        workbook = RubyXL::Parser.parse(filePath)
        workbook[0].sheet_data.rows.each_with_index { |row, i|
            if i != 0
                add_or_update_package(row[0].value, row[1].value, row[2].value, row[3].value)
            end
        }
    end
    
    def add_or_update_package(package_id, item_id, label, value)
        package = @packages.find {|pack| pack.get_id == package_id}
        if (package == nil)
            package = Package.new(package_id)
            @packages << package
        end
        package.add_or_update_item(item_id, label, value)
    end

    def get_packages()
        @packages
    end

    def print()
        printf("Registered Orders:\n")
        @packages.each { |package|
            puts "----------------------------------"
            package.print()
            
        }
    end
end

orders = Orders.new
orders.init_from_xlsx("./Orders.xlsx")
orders.print

db = PG.connect(dbname: 'test', user: 'postgres')
res = db.exec(%{
    INSERT INTO orders
    (orderid, odername)
    VALUES
    (0, 'Order')
    ON CONFLICT (orderid) DO UPDATE SET odername = excluded.odername;
})

orders.get_packages.each { |package|
    res = db.exec(%{
        INSERT INTO packages
        (orderid, packageid)
        VALUES
        (0, #{package.get_id})
        ON CONFLICT (packageid) DO UPDATE SET orderid = excluded.orderid;
    })
    package.get_items.each { |item|
        res = db.exec(%{
            INSERT INTO items
            (packageid, itemid, name, price, ref, warranty, duration)
            VALUES
            (#{package.get_id}, #{(package.get_id + 1) * 1000 + item.get_id}, #{item.get_label("name")}, #{item.get_label("price")}, #{item.get_label("ref")}, #{item.get_label("warranty")}, #{item.get_label("duration")})
            ON CONFLICT (itemid) DO UPDATE SET 
                itemid = excluded.itemid,
                name = excluded.name,
                price = excluded.price,
                "ref" = excluded.ref,
                warranty = excluded.warranty,
                duration = excluded.duration
        })
    }
}
db.close
